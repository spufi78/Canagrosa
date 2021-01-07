VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{77DDF307-D82B-4757-8B3A-106EC9D75D4B}#8.0#0"; "tdbg8.ocx"
Object = "{76C4A5C3-6A01-4523-911A-8FA5928ECD6B}#1.0#0"; "miComboBCA.ocx"
Begin VB.Form frmOfertas_Detalle 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Oferta"
   ClientHeight    =   10155
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9990
   Icon            =   "frmOfertas_Detalle.frx":0000
   KeyPreview      =   -1  'True
   MaxButton       =   0   'False
   NegotiateMenus  =   0   'False
   ScaleHeight     =   10155
   ScaleWidth      =   9990
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdAsignarObra 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Asignar Precios a Obra"
      Height          =   885
      Left            =   2430
      Picture         =   "frmOfertas_Detalle.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   45
      Top             =   9240
      Width           =   2190
   End
   Begin VB.TextBox txtestado 
      Height          =   405
      Left            =   3630
      TabIndex        =   44
      Top             =   9390
      Visible         =   0   'False
      Width           =   1545
   End
   Begin VB.CommandButton cmdImprimir 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Imprimir"
      Height          =   885
      Left            =   60
      Style           =   1  'Graphical
      TabIndex        =   41
      Top             =   9240
      Width           =   1155
   End
   Begin VB.CommandButton cmdCorreo 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Correo"
      Height          =   885
      Index           =   2
      Left            =   1260
      Style           =   1  'Graphical
      TabIndex        =   40
      Top             =   9240
      Width           =   1155
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00C0C0C0&
      Caption         =   " Datos de la Oferta"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   675
      Left            =   60
      TabIndex        =   31
      Top             =   450
      Width           =   9855
      Begin VB.TextBox txtdatos 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H000000FF&
         Height          =   330
         Index           =   0
         Left            =   1350
         Locked          =   -1  'True
         MaxLength       =   75
         TabIndex        =   32
         Top             =   240
         Width           =   1650
      End
      Begin MSComCtl2.DTPicker fecha 
         Height          =   345
         Left            =   3840
         TabIndex        =   35
         Top             =   240
         Width           =   1425
         _ExtentX        =   2514
         _ExtentY        =   609
         _Version        =   393216
         Format          =   16515073
         CurrentDate     =   40679
      End
      Begin MSDataListLib.DataCombo cmbAgente 
         Height          =   315
         Left            =   6300
         TabIndex        =   0
         Top             =   240
         Width           =   3345
         _ExtentX        =   5900
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
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Agente"
         Height          =   195
         Left            =   5580
         TabIndex        =   36
         Top             =   270
         Width           =   510
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Número"
         Height          =   195
         Index           =   8
         Left            =   225
         TabIndex        =   34
         Top             =   285
         Width           =   555
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Fecha"
         Height          =   195
         Index           =   6
         Left            =   3180
         TabIndex        =   33
         Top             =   300
         Width           =   450
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00C0C0C0&
      Caption         =   " Datos de la Obra"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1005
      Left            =   60
      TabIndex        =   28
      Top             =   3840
      Width           =   9855
      Begin VB.TextBox txtdatos 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   315
         Index           =   11
         Left            =   1350
         MaxLength       =   75
         TabIndex        =   11
         Top             =   615
         Width           =   4080
      End
      Begin VB.TextBox txtdatos 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   315
         Index           =   10
         Left            =   1350
         MaxLength       =   55
         TabIndex        =   10
         Top             =   270
         Width           =   8280
      End
      Begin MSDataListLib.DataCombo cmbTipoObra 
         Height          =   315
         Left            =   7020
         TabIndex        =   12
         Top             =   630
         Width           =   2625
         _ExtentX        =   4630
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
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Tipo de Obra"
         Height          =   195
         Left            =   5820
         TabIndex        =   43
         Top             =   690
         Width           =   930
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Población"
         Height          =   195
         Index           =   13
         Left            =   225
         TabIndex        =   30
         Top             =   660
         Width           =   705
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Domicilio"
         Height          =   195
         Index           =   12
         Left            =   225
         TabIndex        =   29
         Top             =   315
         Width           =   630
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0C0C0&
      Caption         =   " Datos de la Oferta"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2685
      Left            =   60
      TabIndex        =   17
      Top             =   1140
      Width           =   9855
      Begin VB.TextBox txtdatos 
         Appearance      =   0  'Flat
         Height          =   315
         Index           =   9
         Left            =   5910
         MaxLength       =   50
         TabIndex        =   8
         Top             =   1950
         Width           =   3720
      End
      Begin VB.TextBox txtdatos 
         Appearance      =   0  'Flat
         Height          =   315
         Index           =   5
         Left            =   1350
         MaxLength       =   75
         TabIndex        =   7
         Top             =   1950
         Width           =   3450
      End
      Begin VB.TextBox txtdatos 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   315
         Index           =   4
         Left            =   1350
         MaxLength       =   75
         TabIndex        =   4
         Top             =   1620
         Width           =   1995
      End
      Begin VB.TextBox txtdatos 
         Appearance      =   0  'Flat
         Height          =   315
         Index           =   3
         Left            =   1350
         MaxLength       =   55
         TabIndex        =   3
         Top             =   1290
         Width           =   8280
      End
      Begin VB.TextBox txtdatos 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   315
         Index           =   1
         Left            =   1350
         MaxLength       =   75
         TabIndex        =   1
         Top             =   600
         Width           =   8280
      End
      Begin VB.TextBox txtdatos 
         Appearance      =   0  'Flat
         Height          =   315
         Index           =   2
         Left            =   1350
         MaxLength       =   55
         TabIndex        =   2
         Top             =   945
         Width           =   8280
      End
      Begin VB.TextBox txtdatos 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   315
         Index           =   6
         Left            =   4590
         MaxLength       =   50
         TabIndex        =   5
         Top             =   1620
         Width           =   2025
      End
      Begin VB.TextBox txtdatos 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   315
         Index           =   7
         Left            =   7650
         MaxLength       =   30
         TabIndex        =   6
         Top             =   1620
         Width           =   1980
      End
      Begin VB.TextBox txtdatos 
         Appearance      =   0  'Flat
         Height          =   315
         Index           =   8
         Left            =   1350
         MaxLength       =   50
         TabIndex        =   9
         Top             =   2280
         Width           =   8280
      End
      Begin vb6projectpryComboBCA.miComboBCA cmbCliente 
         Height          =   375
         Left            =   1350
         TabIndex        =   42
         Top             =   240
         Width           =   8265
         _ExtentX        =   14579
         _ExtentY        =   661
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Cliente"
         Height          =   195
         Index           =   10
         Left            =   210
         TabIndex        =   37
         Top             =   300
         Width           =   480
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Forma Pago"
         Height          =   195
         Index           =   4
         Left            =   4950
         TabIndex        =   27
         Top             =   2010
         Width           =   855
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "A. Atención"
         Height          =   195
         Index           =   3
         Left            =   210
         TabIndex        =   26
         Top             =   1980
         Width           =   825
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "N.I.F."
         Height          =   195
         Index           =   2
         Left            =   210
         TabIndex        =   25
         Top             =   1680
         Width           =   390
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Nombre"
         Height          =   195
         Index           =   0
         Left            =   210
         TabIndex        =   23
         Top             =   645
         Width           =   555
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Direccion"
         Height          =   195
         Index           =   1
         Left            =   210
         TabIndex        =   22
         Top             =   990
         Width           =   675
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Telefono"
         Height          =   195
         Index           =   11
         Left            =   3735
         TabIndex        =   21
         Top             =   1680
         Width           =   630
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Fax"
         Height          =   195
         Index           =   15
         Left            =   7020
         TabIndex        =   20
         Top             =   1680
         Width           =   255
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Población"
         Height          =   195
         Index           =   7
         Left            =   210
         TabIndex        =   19
         Top             =   1320
         Width           =   705
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "e-Mail"
         Height          =   195
         Index           =   9
         Left            =   210
         TabIndex        =   18
         Top             =   2340
         Width           =   420
      End
   End
   Begin VB.CommandButton cmdSalir 
      BackColor       =   &H00E0E0E0&
      Caption         =   "ESC-Salir"
      Height          =   885
      Left            =   8790
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   9240
      Width           =   1155
   End
   Begin VB.CommandButton cmdborrar 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Borrar &Línea"
      Height          =   885
      Left            =   5940
      Picture         =   "frmOfertas_Detalle.frx":0BD4
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   9210
      Visible         =   0   'False
      Width           =   1065
   End
   Begin TrueDBGrid80.TDBGrid gridTarifa 
      Height          =   4335
      Left            =   60
      TabIndex        =   13
      Top             =   4860
      Width           =   9885
      _ExtentX        =   17436
      _ExtentY        =   7646
      _LayoutType     =   4
      _RowHeight      =   -2147483647
      _WasPersistedAsPixels=   0
      Columns(0)._VlistStyle=   0
      Columns(0)._MaxComboItems=   5
      Columns(0).Caption=   "Ref."
      Columns(0).DataField=   ""
      Columns(0).DropDown=   "tArticulos"
      Columns(0).DropDown.vt=   8
      Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(1)._VlistStyle=   0
      Columns(1)._MaxComboItems=   5
      Columns(1).Caption=   "Material"
      Columns(1).DataField=   ""
      Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(2)._VlistStyle=   0
      Columns(2)._MaxComboItems=   5
      Columns(2).Caption=   "Precio en Fabrica"
      Columns(2).DataField=   ""
      Columns(2).NumberFormat=   "Currency"
      Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(3)._VlistStyle=   0
      Columns(3)._MaxComboItems=   5
      Columns(3).Caption=   "Precio en Obra"
      Columns(3).DataField=   ""
      Columns(3).NumberFormat=   "Currency"
      Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns.Count   =   4
      Splits(0)._UserFlags=   0
      Splits(0).MarqueeStyle=   1
      Splits(0).AllowRowSizing=   0   'False
      Splits(0).RecordSelectors=   0   'False
      Splits(0).RecordSelectorWidth=   503
      Splits(0).AllowColSelect=   0   'False
      Splits(0).DividerColor=   12632256
      Splits(0).SpringMode=   0   'False
      Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
      Splits(0)._ColumnProps(0)=   "Columns.Count=4"
      Splits(0)._ColumnProps(1)=   "Column(0).Width=1640"
      Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
      Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=1561"
      Splits(0)._ColumnProps(4)=   "Column(0)._EditAlways=0"
      Splits(0)._ColumnProps(5)=   "Column(0)._ColStyle=1"
      Splits(0)._ColumnProps(6)=   "Column(0).AllowFocus=0"
      Splits(0)._ColumnProps(7)=   "Column(0).Order=1"
      Splits(0)._ColumnProps(8)=   "Column(0).AutoDropDown=1"
      Splits(0)._ColumnProps(9)=   "Column(0).DropDownList=1"
      Splits(0)._ColumnProps(10)=   "Column(0).AutoCompletion=1"
      Splits(0)._ColumnProps(11)=   "Column(1).Width=8864"
      Splits(0)._ColumnProps(12)=   "Column(1).DividerColor=0"
      Splits(0)._ColumnProps(13)=   "Column(1)._WidthInPix=8784"
      Splits(0)._ColumnProps(14)=   "Column(1)._EditAlways=0"
      Splits(0)._ColumnProps(15)=   "Column(1)._ColStyle=8196"
      Splits(0)._ColumnProps(16)=   "Column(1).AllowFocus=0"
      Splits(0)._ColumnProps(17)=   "Column(1).Order=2"
      Splits(0)._ColumnProps(18)=   "Column(2).Width=3175"
      Splits(0)._ColumnProps(19)=   "Column(2).DividerColor=0"
      Splits(0)._ColumnProps(20)=   "Column(2)._WidthInPix=3096"
      Splits(0)._ColumnProps(21)=   "Column(2)._EditAlways=0"
      Splits(0)._ColumnProps(22)=   "Column(2)._ColStyle=2"
      Splits(0)._ColumnProps(23)=   "Column(2).Order=3"
      Splits(0)._ColumnProps(24)=   "Column(3).Width=3016"
      Splits(0)._ColumnProps(25)=   "Column(3).DividerColor=0"
      Splits(0)._ColumnProps(26)=   "Column(3)._WidthInPix=2937"
      Splits(0)._ColumnProps(27)=   "Column(3)._EditAlways=0"
      Splits(0)._ColumnProps(28)=   "Column(3)._ColStyle=2"
      Splits(0)._ColumnProps(29)=   "Column(3).Order=4"
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
      _StyleDefs(36)  =   "Splits(0).Columns(0).Style:id=24,.parent=11,.alignment=2,.bold=0,.fontsize=975"
      _StyleDefs(37)  =   ":id=24,.italic=0,.underline=0,.strikethrough=0,.charset=0"
      _StyleDefs(38)  =   ":id=24,.fontname=MS Sans Serif"
      _StyleDefs(39)  =   "Splits(0).Columns(0).HeadingStyle:id=21,.parent=12"
      _StyleDefs(40)  =   "Splits(0).Columns(0).FooterStyle:id=22,.parent=13"
      _StyleDefs(41)  =   "Splits(0).Columns(0).EditorStyle:id=23,.parent=15,.bold=0,.fontsize=975"
      _StyleDefs(42)  =   ":id=23,.italic=0,.underline=0,.strikethrough=0,.charset=0"
      _StyleDefs(43)  =   ":id=23,.fontname=MS Sans Serif"
      _StyleDefs(44)  =   "Splits(0).Columns(1).Style:id=28,.parent=11,.locked=-1"
      _StyleDefs(45)  =   "Splits(0).Columns(1).HeadingStyle:id=25,.parent=12"
      _StyleDefs(46)  =   "Splits(0).Columns(1).FooterStyle:id=26,.parent=13"
      _StyleDefs(47)  =   "Splits(0).Columns(1).EditorStyle:id=27,.parent=15"
      _StyleDefs(48)  =   "Splits(0).Columns(2).Style:id=32,.parent=11,.alignment=1"
      _StyleDefs(49)  =   "Splits(0).Columns(2).HeadingStyle:id=29,.parent=12"
      _StyleDefs(50)  =   "Splits(0).Columns(2).FooterStyle:id=30,.parent=13"
      _StyleDefs(51)  =   "Splits(0).Columns(2).EditorStyle:id=31,.parent=15"
      _StyleDefs(52)  =   "Splits(0).Columns(3).Style:id=36,.parent=11,.alignment=1"
      _StyleDefs(53)  =   "Splits(0).Columns(3).HeadingStyle:id=33,.parent=12"
      _StyleDefs(54)  =   "Splits(0).Columns(3).FooterStyle:id=34,.parent=13"
      _StyleDefs(55)  =   "Splits(0).Columns(3).EditorStyle:id=35,.parent=15"
      _StyleDefs(56)  =   "Named:id=37:Normal"
      _StyleDefs(57)  =   ":id=37,.parent=0,.alignment=3,.bgcolor=&H80000018&,.bold=0,.fontsize=825"
      _StyleDefs(58)  =   ":id=37,.italic=0,.underline=0,.strikethrough=0,.charset=0"
      _StyleDefs(59)  =   ":id=37,.fontname=MS Sans Serif"
      _StyleDefs(60)  =   "Named:id=38:Heading"
      _StyleDefs(61)  =   ":id=38,.parent=37,.valignment=2,.bgcolor=&HC0C0C0&,.fgcolor=&H80000012&"
      _StyleDefs(62)  =   ":id=38,.wraptext=-1"
      _StyleDefs(63)  =   "Named:id=39:Footing"
      _StyleDefs(64)  =   ":id=39,.parent=37,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(65)  =   "Named:id=40:Selected"
      _StyleDefs(66)  =   ":id=40,.parent=37,.bgcolor=&H8080FF&,.fgcolor=&H0&,.bold=0,.fontsize=825"
      _StyleDefs(67)  =   ":id=40,.italic=0,.underline=0,.strikethrough=0,.charset=0"
      _StyleDefs(68)  =   ":id=40,.fontname=MS Sans Serif"
      _StyleDefs(69)  =   "Named:id=41:Caption"
      _StyleDefs(70)  =   ":id=41,.parent=38,.alignment=2"
      _StyleDefs(71)  =   "Named:id=42:HighlightRow"
      _StyleDefs(72)  =   ":id=42,.parent=37,.bgcolor=&H80000008&,.fgcolor=&H80000005&"
      _StyleDefs(73)  =   "Named:id=43:EvenRow"
      _StyleDefs(74)  =   ":id=43,.parent=37,.bgcolor=&HFFFFFF&"
      _StyleDefs(75)  =   "Named:id=44:OddRow"
      _StyleDefs(76)  =   ":id=44,.parent=37,.bgcolor=&HE9E9E9&"
      _StyleDefs(77)  =   "Named:id=47:RecordSelector"
      _StyleDefs(78)  =   ":id=47,.parent=38"
      _StyleDefs(79)  =   "Named:id=50:FilterBar"
      _StyleDefs(80)  =   ":id=50,.parent=37"
   End
   Begin MSDataListLib.DataCombo cmbEstado 
      Height          =   360
      Left            =   6570
      TabIndex        =   38
      Top             =   30
      Width           =   3345
      _ExtentX        =   5900
      _ExtentY        =   635
      _Version        =   393216
      Appearance      =   0
      Style           =   2
      ListField       =   ""
      BoundColumn     =   ""
      Text            =   ""
      Object.DataMember      =   ""
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.CommandButton cmdAceptar 
      BackColor       =   &H00E0E0E0&
      Caption         =   "F10-Aceptar"
      Height          =   885
      Left            =   7590
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   9240
      Width           =   1155
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0FFFF&
      Caption         =   "ESTADO"
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
      Left            =   5550
      TabIndex        =   39
      Top             =   90
      Width           =   945
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      Caption         =   "  Detalle de Oferta"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   420
      Index           =   5
      Left            =   0
      TabIndex        =   24
      Top             =   0
      Width           =   10320
   End
End
Attribute VB_Name = "frmOfertas_Detalle"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public pk As Long

Dim xTarifa As New XArrayDB

Const filasTarifa As Integer = 50
Const ColTarifa As Integer = 4
Private Enum ColsTarifa
    ID = 0
    ARTICULO = 1
    PFABRICA = 2
    POBRA = 3
End Enum

Private Sub cmbCliente_change()
   On Error GoTo cmbCliente_change_Error

    If cmbCliente.getTEXTO <> "" And pk = 0 Then
        ' Copiar los datos del cliente
        If MsgBox("¿Desea informar los datos de la oferta con los del cliente?", vbQuestion + vbYesNo, App.Title) = vbYes Then
            Dim ocliente As New clsCliente
            With ocliente
                .CargaCliente cmbCliente.getPK_SALIDA
                txtDatos(1) = .getNOMBRE
                txtDatos(2) = .getDIRECCION
                Dim oMunicipio As New clsMunicipios
                If oMunicipio.Cargar(.getMUNICIPIO_ID) Then
                    txtDatos(3) = oMunicipio.getNOMBRE
                End If
                txtDatos(4) = .getCIF
                
                txtDatos(6) = .getTELEFONO
                txtDatos(7) = .getFAX
                
                txtDatos(5) = .getRAZON
                txtDatos(8) = .getEMAIL
                Dim oFp As New clsForma_pago
                oFp.Cargar .getFORMA_PAGO
                txtDatos(9) = oFp.getNOMBRE
            End With
            Set ocliente = Nothing
            txtDatos(10).SetFocus
        End If
    End If

   On Error GoTo 0
   Exit Sub

cmbCliente_change_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure cmbCliente_change of Formulario frmOfertas_Detalle"
    
End Sub

Private Sub cmdAceptar_Click()
   On Error GoTo cmdAceptar_Click_Error

    If validar Then
        Dim oOferta As New clsOfertas
        With oOferta
            If cmbCliente.getTEXTO = "" Then
                .setCLIENTE_ID = 0
            Else
                .setCLIENTE_ID = cmbCliente.getPK_SALIDA
            End If
            .setNOMBRE = txtDatos(1)
            .setDIRECCION = txtDatos(2)
            .setPOBLACION = txtDatos(3)
            .setNIF = txtDatos(4)
            .setAA = txtDatos(5)
            .setTELEFONO = txtDatos(6)
            .setFAX = txtDatos(7)
            .setEMAIL = txtDatos(8)
            .setFORMA_PAGO = txtDatos(9)
            .setOBRA_DOMICILIO = txtDatos(10)
            .setOBRA_POBLACION = txtDatos(11)
            .setESTADO_ID = cmbEstado.BoundText
            .setFECHA = Format(fecha, "yyyy-mm-dd")
            .setAGENTE_ID = cmbAgente.BoundText
            .setHORA = Format(Time, "HH:MM:SS")
            If cmbTipoObra.Text = "" Then
                .setOBRA_TIPO = 0
            Else
                .setOBRA_TIPO = cmbTipoObra.BoundText
            End If
            If pk = 0 Then
                pk = .Insertar
                If pk = 0 Then
                    Exit Sub
                End If
            Else
                If Not .Modificar(pk) Then
                    Exit Sub
                End If
            End If
            
            ' Detalle
            Dim oDO As New clsOfertas_detalle
            Dim i As Integer
            oDO.Eliminar pk
            For i = xTarifa.LowerBound(1) To xTarifa.UpperBound(1)
                If Trim(xTarifa.Value(i, ColsTarifa.ID)) <> "" Then
                    With oDO
                        .setOFERTA_ID = pk
                        .setORDEN = i
                        .setMATERIAL_ID = Trim(xTarifa.Value(i, ColsTarifa.ID))
                        If Trim(xTarifa.Value(i, ColsTarifa.PFABRICA)) = "" Then
                            .setPRECIO_FABRICA = moneda_bd("0")
                        Else
                            .setPRECIO_FABRICA = moneda_bd(Trim(xTarifa.Value(i, ColsTarifa.PFABRICA)))
                        End If
                        If Trim(xTarifa.Value(i, ColsTarifa.POBRA)) = "" Then
                            .setPRECIO_OBRA = moneda_bd("0")
                        Else
                            .setPRECIO_OBRA = moneda_bd(Trim(xTarifa.Value(i, ColsTarifa.POBRA)))
                        End If
                        If .Insertar = 0 Then
                            Exit Sub
                        Else
                        End If
                    End With
                End If
            Next
            Set oDO = Nothing
        End With
        ' Si se acepta la oferta, el estado anterior era distinto, y es nuevo cliente, monstramos la pantalla de cliente
        If cmbEstado.BoundText = ENUM_OFERTAS_ESTADOS.OFERTAR_ESTADOS_ACEPTADA And _
            txtestado <> ENUM_OFERTAS_ESTADOS.OFERTAR_ESTADOS_ACEPTADA Then
            If cmbCliente.getPK_SALIDA = 0 Then
                If MsgBox("¿Desea dar de alta el cliente?", vbQuestion + vbYesNo, App.Title) = vbYes Then
                    With frmClientes
                        .txtDatos(1) = txtDatos(1) ' Nombre
                        .txtDatos(2) = txtDatos(2) ' Direccion
                        .txtDatos(8) = txtDatos(4) ' NIF
                        .txtDatos(6) = txtDatos(6) ' Telefono
                        .txtDatos(0) = txtDatos(7) ' Fax
                        .txtDatos(10) = txtDatos(5) ' A la atencion
                        .txtDatos(16) = txtDatos(8) ' Email
                        .Show 1
                    End With
                End If
            End If
        End If
        Set oOferta = Nothing
        MsgBox "La Oferta se ha almacenado correctamente.", vbInformation, App.Title
        Unload Me
    End If

   On Error GoTo 0
   Exit Sub

cmdAceptar_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure cmdAceptar_Click of Formulario frmOfertas_Detalle"
End Sub

Private Sub cmdAsignarObra_Click()
    If cmbEstado.BoundText <> ENUM_OFERTAS_ESTADOS.OFERTAR_ESTADOS_ACEPTADA Then
        MsgBox "La oferta debe estar 'ACEPTADA' para asignarla a una obra.", vbInformation, App.Title
    Else
        frmOfertas_Asignar.pk = pk
        frmOfertas_Asignar.Show 1
    End If
End Sub

Private Sub cmdBorrar_Click()
    On Error Resume Next
    Dim i As Integer
    For i = 0 To ColTarifa
        gridTarifa.SelBookmarks.Add gridTarifa.Bookmark
        xTarifa(gridTarifa.Bookmark, i) = ""
        gridTarifa.SelBookmarks.Remove 0
    Next
    gridTarifa.Refresh
    gridTarifa.SetFocus
End Sub

Private Sub cmdCorreo_Click(Index As Integer)
    If pk > 0 Then
        Dim oOferta As New clsOfertas
        oOferta.Correo pk, False, True, 1
        Set oOferta = Nothing
    End If
End Sub

Private Sub cmdImprimir_Click()
    If pk > 0 Then
        Dim oOferta As New clsOfertas
        oOferta.imprimir pk, False, True, 1, ""
        Set oOferta = Nothing
    End If
End Sub

Private Sub cmdSalir_Click()
'    If MsgBox("¿Esta seguro que desea salir?", vbYesNo + vbQuestion, App.Title) = vbYes Then
        Unload Me
'    End If
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
    cargar_botones Me
    cargar_combos
    inicializar_grid
    If pk > 0 Then
        Cargar
    Else
        cmbEstado.BoundText = ENUM_OFERTAS_ESTADOS.OFERTAR_ESTADOS_PENDIENTE
        cmbEstado.Locked = True
        Dim oOferta As New clsOfertas
        oOferta.CrearID
        txtDatos(0) = oOferta.getID_OFERTA
        fecha = Date
        cargar_articulos
'        cmbAgente.BoundText = usuario.getID_EMPLEADO
    End If
    txtestado = cmbEstado.BoundText
End Sub

Private Sub inicializar_grid()
   On Error GoTo inicializar_grid_Error
   
    gridTarifa.Col = 0
    gridTarifa.Row = 0
    xTarifa.Clear
    xTarifa.ReDim 0, filasTarifa, 0, ColTarifa
    xTarifa.Clear
    Set gridTarifa.Array = xTarifa
    gridTarifa.Refresh

   On Error GoTo 0
   Exit Sub

inicializar_grid_Error:
    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure inicializar_grid of Formulario frmCE_Recepcion_Nuevo2"
End Sub

Private Sub Cargar()
    On Error GoTo fallo
    Dim oOferta As New clsOfertas
    If oOferta.Carga(pk) = True Then
        cmbEstado.BoundText = oOferta.getESTADO_ID
        txtDatos(0) = oOferta.getNUMERO & "/" & Year(oOferta.getFECHA)
        fecha = oOferta.getFECHA
        cmbAgente.BoundText = oOferta.getAGENTE_ID
        cmbCliente.MostrarElemento oOferta.getCLIENTE_ID
        txtDatos(1) = oOferta.getNOMBRE
        txtDatos(2) = oOferta.getDIRECCION
        txtDatos(3) = oOferta.getPOBLACION
        txtDatos(4) = oOferta.getNIF
        txtDatos(5) = oOferta.getAA
        txtDatos(6) = oOferta.getTELEFONO
        txtDatos(7) = oOferta.getFAX
        txtDatos(8) = oOferta.getEMAIL
        txtDatos(9) = oOferta.getFORMA_PAGO
        txtDatos(10) = oOferta.getOBRA_DOMICILIO
        txtDatos(11) = oOferta.getOBRA_POBLACION
        cmbTipoObra.BoundText = oOferta.getOBRA_TIPO
       ' Cargamos los datos de la tarifa
       Dim oOD As New clsOfertas_detalle
       Dim rs As ADODB.Recordset
       Set rs = oOD.Listado(pk)
       If rs.RecordCount > 0 Then
            Dim fila As Long
            fila = 0
            Do
                xTarifa(fila, ColsTarifa.ID) = CStr(rs(0))
                xTarifa(fila, ColsTarifa.ARTICULO) = CStr(rs(1))
                xTarifa(fila, ColsTarifa.PFABRICA) = CStr(rs(2))
                xTarifa(fila, ColsTarifa.POBRA) = CStr(rs(3))
                rs.MoveNext
                fila = fila + 1
            Loop Until rs.EOF
            gridTarifa.Row = 0
            gridTarifa.Col = 0
            gridTarifa.Refresh
'            gridTarifa.SetFocus
        End If
        Set oOD = Nothing
        Set rs = Nothing
    Else
        MsgBox "Error al cargar la tarifa.", vbInformation, App.Title
    End If
    Set oOferta = Nothing
    Exit Sub
fallo:
    MsgBox "Error al cargar los datos de la oferta. " & Err.Description, vbCritical, App.Title
End Sub

Private Sub gridTarifa_KeyPress(KeyAscii As Integer)
    If (gridTarifa.Col = ColsTarifa.PFABRICA Or gridTarifa.Col = ColsTarifa.POBRA) And KeyAscii = 46 Then
         KeyAscii = 44
    End If
End Sub

Private Sub txtDatos_GotFocus(Index As Integer)
    txtDatos(Index).BackColor = &HC0E0FF
    txtDatos(Index).SelStart = 0
    txtDatos(Index).SelLength = Len(txtDatos(Index))
End Sub

Private Sub txtDatos_LostFocus(Index As Integer)
    txtDatos(Index).BackColor = vbWhite
End Sub

Private Sub cargar_combos()
    llenar_combo cmbCliente, New clsCliente, 0, frmClientes, " ANULADO = 0 "
    Cargar_Combo cmbAgente, New clsComercial
    Dim oD As New clsDecodificadora
    oD.Cargar_Combo cmbEstado, DECODIFICADORA.D_OFERTAS_ESTADOS
    oD.Cargar_Combo cmbTipoObra, DECODIFICADORA.D_TIPOS_OBRAS
End Sub

Private Function validar() As Boolean
    validar = True
    If txtDatos(1) = "" Then
        MsgBox "El nombre del cliente debe estar informado.", vbExclamation, App.Title
        txtDatos(1).SetFocus
        validar = False
    End If
'    If txtdatos(4) = "" Then
'        MsgBox "El NIF del cliente debe estar informado.", vbExclamation, App.Title
'        txtdatos(4).SetFocus
'        validar = False
'    End If
'    If txtdatos(6) = "" Then
'        MsgBox "El teléfono del cliente debe estar informado.", vbExclamation, App.Title
'        txtdatos(6).SetFocus
'        validar = False
'    End If
'    If txtdatos(5) = "" Then
'        MsgBox "El campo 'A la atención' debe estar informado.", vbExclamation, App.Title
'        txtdatos(5).SetFocus
'        validar = False
'    End If
'    If cmbAgente.Text = "" Then
'        MsgBox "Debe informar el Agente.", vbExclamation, App.Title
'        cmbAgente.SetFocus
'        validar = False
'    End If
    If cmbTipoObra.Text = "" Then
        MsgBox "Indique el tipo de la obra.", vbExclamation, App.Title
        cmbTipoObra.SetFocus
        validar = False
        Exit Function
    End If
End Function

Private Sub cargar_articulos()
       Dim consulta As String
       Dim rs As ADODB.Recordset
       consulta = "SELECT B.VALOR, B.DESCRIPCION " & _
                   "  FROM DECODIFICADORA B " & _
                   " WHERE B.CODIGO = " & DECODIFICADORA.D_OFERTAS_MATERIALES & _
                   "   AND B.DESCRIPCION <> '' " & _
                   " ORDER BY B.PARAMETROS "
       Set rs = datos_bd(consulta)
       If rs.RecordCount > 0 Then
            Dim fila As Long
            fila = 0
            Do
                xTarifa(fila, ColsTarifa.ID) = CStr(rs(0))
                xTarifa(fila, ColsTarifa.ARTICULO) = CStr(rs(1))
                xTarifa(fila, ColsTarifa.PFABRICA) = CStr(moneda("0"))
                xTarifa(fila, ColsTarifa.POBRA) = CStr(moneda("0"))
                rs.MoveNext
                fila = fila + 1
            Loop Until rs.EOF
            gridTarifa.Row = 0
            gridTarifa.Col = 0
            gridTarifa.Refresh
'            gridTarifa.SetFocus
        End If
End Sub
