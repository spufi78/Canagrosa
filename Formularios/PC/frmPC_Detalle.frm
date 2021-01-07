VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{CEA3B7F6-2847-4E5E-A551-DB7A62489D44}#46.0#0"; "miCombo.ocx"
Object = "{77DDF307-D82B-4757-8B3A-106EC9D75D4B}#8.0#0"; "tdbg8.ocx"
Begin VB.Form frmPC_Detalle 
   BackColor       =   &H00C0C0C0&
   Caption         =   "Gestión de pedido de Producto Controlado"
   ClientHeight    =   12435
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   13500
   Icon            =   "frmPC_Detalle.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   12435
   ScaleWidth      =   13500
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame4 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   "CoC/Informe de ensayo del Fabricante"
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
      Height          =   870
      Left            =   6615
      TabIndex        =   42
      Top             =   8190
      Width           =   6630
      Begin VB.TextBox txtInformeFabricante 
         Appearance      =   0  'Flat
         Height          =   645
         Left            =   45
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   43
         Top             =   180
         Width           =   6495
      End
   End
   Begin VB.Frame Frame3 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   "Identificación"
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
      Height          =   870
      Left            =   45
      TabIndex        =   40
      Top             =   8190
      Width           =   6540
      Begin VB.TextBox txtIdentificacion 
         Appearance      =   0  'Flat
         Height          =   645
         Left            =   45
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   41
         Top             =   180
         Width           =   6405
      End
   End
   Begin VB.TextBox txtDOC_ID 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   8190
      TabIndex        =   28
      Top             =   11835
      Visible         =   0   'False
      Width           =   1680
   End
   Begin VB.CommandButton cmdFactura 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Factura"
      Height          =   870
      Left            =   10080
      Style           =   1  'Graphical
      TabIndex        =   27
      Top             =   11520
      Visible         =   0   'False
      Width           =   1050
   End
   Begin VB.TextBox txtcodigo 
      Height          =   375
      Left            =   2100
      TabIndex        =   24
      Top             =   11940
      Visible         =   0   'False
      Width           =   705
   End
   Begin VB.Frame Frame2 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   "Datos del Cliente"
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
      Height          =   2190
      Left            =   6690
      TabIndex        =   19
      Top             =   540
      Width           =   6795
      Begin VB.TextBox txtDatos 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H000000FF&
         Height          =   315
         Index           =   1
         Left            =   1275
         TabIndex        =   4
         Top             =   975
         Width           =   5115
      End
      Begin MSComCtl2.DTPicker fechaCliente 
         Height          =   330
         Left            =   1260
         TabIndex        =   5
         Top             =   1350
         Width           =   1425
         _ExtentX        =   2514
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
         Format          =   60751873
         CurrentDate     =   38000
         MinDate         =   2
      End
      Begin pryCombo.miCombo cmbCliente 
         Height          =   345
         Left            =   1275
         TabIndex        =   3
         Top             =   270
         Width           =   5460
         _ExtentX        =   9631
         _ExtentY        =   609
      End
      Begin pryCombo.miCombo cmbPedidos 
         Height          =   330
         Left            =   1275
         TabIndex        =   31
         Top             =   630
         Width           =   5460
         _ExtentX        =   9631
         _ExtentY        =   582
      End
      Begin VB.Label Label2 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Pedido"
         ForeColor       =   &H00000000&
         Height          =   285
         Index           =   24
         Left            =   120
         TabIndex        =   32
         Top             =   675
         Width           =   690
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Cliente"
         Height          =   195
         Index           =   0
         Left            =   120
         TabIndex        =   23
         Top             =   330
         Width           =   480
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Fecha Pedido"
         Height          =   195
         Index           =   4
         Left            =   120
         TabIndex        =   21
         Top             =   1395
         Width           =   990
      End
      Begin VB.Label lblCampos 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Orden Cliente"
         Height          =   240
         Index           =   5
         Left            =   120
         TabIndex        =   20
         Top             =   1035
         Width           =   1005
      End
   End
   Begin VB.CommandButton cmdElimina 
      BackColor       =   &H00E0E0E0&
      Height          =   690
      Left            =   12510
      Picture         =   "frmPC_Detalle.frx":08CA
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   6645
      Width           =   675
   End
   Begin VB.CommandButton cmdcancel 
      BackColor       =   &H00E0E0E0&
      Caption         =   "ESC-Salir"
      Height          =   870
      Left            =   12225
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   11520
      Width           =   1050
   End
   Begin VB.CommandButton cmdok 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Aceptar"
      Height          =   870
      Left            =   11145
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   11520
      Width           =   1050
   End
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   "Datos del Pedido"
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
      Height          =   2190
      Left            =   30
      TabIndex        =   12
      Top             =   540
      Width           =   6615
      Begin VB.CheckBox chkFechaCaducidadNA 
         BackColor       =   &H00C0C0C0&
         Caption         =   "No Aplica"
         Height          =   195
         Left            =   2835
         TabIndex        =   39
         Top             =   1845
         Value           =   1  'Checked
         Width           =   1320
      End
      Begin VB.TextBox txtNumCertificado 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   4860
         TabIndex        =   33
         Top             =   1035
         Width           =   1410
      End
      Begin VB.TextBox txtPrecio 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   1305
         TabIndex        =   29
         Top             =   990
         Width           =   1365
      End
      Begin VB.TextBox txtDatos 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H000000FF&
         Height          =   315
         Index           =   0
         Left            =   1305
         TabIndex        =   1
         Top             =   630
         Width           =   2505
      End
      Begin MSComCtl2.DTPicker fechaPedido 
         Height          =   330
         Left            =   4875
         TabIndex        =   2
         Top             =   630
         Width           =   1425
         _ExtentX        =   2514
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
         Format          =   60751873
         CurrentDate     =   38000
         MinDate         =   2
      End
      Begin pryCombo.miCombo cmbProducto 
         Height          =   330
         Left            =   1305
         TabIndex        =   0
         Top             =   270
         Width           =   5190
         _ExtentX        =   9155
         _ExtentY        =   582
      End
      Begin pryCombo.miCombo cmbNormativa 
         Height          =   330
         Left            =   1305
         TabIndex        =   35
         Top             =   1395
         Width           =   5190
         _ExtentX        =   9155
         _ExtentY        =   582
      End
      Begin MSComCtl2.DTPicker fCaducidad 
         Height          =   330
         Left            =   1320
         TabIndex        =   37
         Top             =   1755
         Width           =   1380
         _ExtentX        =   2434
         _ExtentY        =   582
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
         Format          =   60751873
         CurrentDate     =   38000
         MinDate         =   2
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "F.Caducidad"
         Height          =   195
         Index           =   1
         Left            =   135
         TabIndex        =   38
         Top             =   1800
         Width           =   900
      End
      Begin VB.Label lblCampos 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Normativa"
         Height          =   240
         Index           =   3
         Left            =   135
         TabIndex        =   36
         Top             =   1440
         Width           =   945
      End
      Begin VB.Label lblCampos 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Nº Certificado"
         Height          =   240
         Index           =   1
         Left            =   3825
         TabIndex        =   34
         Top             =   1080
         Width           =   1005
      End
      Begin VB.Label lblCampos 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Precio Bote"
         Height          =   240
         Index           =   4
         Left            =   135
         TabIndex        =   30
         Top             =   1035
         Width           =   1005
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Fecha"
         Height          =   195
         Index           =   2
         Left            =   4185
         TabIndex        =   22
         Top             =   675
         Width           =   450
      End
      Begin VB.Label lblCampos 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Producto"
         Height          =   240
         Index           =   0
         Left            =   135
         TabIndex        =   14
         Top             =   315
         Width           =   945
      End
      Begin VB.Label lblCampos 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Número Lote"
         Height          =   240
         Index           =   2
         Left            =   135
         TabIndex        =   13
         Top             =   660
         Width           =   1005
      End
   End
   Begin MSComctlLib.ListView listaReactivos 
      Height          =   2325
      Left            =   60
      TabIndex        =   6
      Top             =   3000
      Width           =   12420
      _ExtentX        =   21908
      _ExtentY        =   4101
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      AllowReorder    =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   12640511
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
   Begin MSComctlLib.ListView listaPedido 
      Height          =   2580
      Left            =   60
      TabIndex        =   8
      Top             =   5580
      Width           =   12405
      _ExtentX        =   21881
      _ExtentY        =   4551
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      AllowReorder    =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   12640511
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
   Begin VB.CommandButton cmdInserta 
      BackColor       =   &H00E0E0E0&
      Height          =   690
      Left            =   12540
      Picture         =   "frmPC_Detalle.frx":1194
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   3900
      Width           =   675
   End
   Begin TrueDBGrid80.TDBGrid grid 
      Height          =   2130
      Left            =   60
      TabIndex        =   25
      Top             =   9315
      Width           =   13170
      _ExtentX        =   23230
      _ExtentY        =   3757
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
      Columns(1).Caption=   "Requisito"
      Columns(1).DataField=   ""
      Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(2)._VlistStyle=   0
      Columns(2)._MaxComboItems=   5
      Columns(2).Caption=   "Valor Parámetro"
      Columns(2).DataField=   ""
      Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(3)._VlistStyle=   0
      Columns(3)._MaxComboItems=   5
      Columns(3).Caption=   "Unidad"
      Columns(3).DataField=   ""
      Columns(3).DropDown=   "tMetodos"
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
      Splits(0)._ColumnProps(1)=   "Column(0).Width=6588"
      Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
      Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=6509"
      Splits(0)._ColumnProps(4)=   "Column(0)._EditAlways=0"
      Splits(0)._ColumnProps(5)=   "Column(0)._ColStyle=1"
      Splits(0)._ColumnProps(6)=   "Column(0).Order=1"
      Splits(0)._ColumnProps(7)=   "Column(1).Width=8123"
      Splits(0)._ColumnProps(8)=   "Column(1).DividerColor=0"
      Splits(0)._ColumnProps(9)=   "Column(1)._WidthInPix=8043"
      Splits(0)._ColumnProps(10)=   "Column(1)._EditAlways=0"
      Splits(0)._ColumnProps(11)=   "Column(1).Order=2"
      Splits(0)._ColumnProps(12)=   "Column(2).Width=6138"
      Splits(0)._ColumnProps(13)=   "Column(2).DividerColor=0"
      Splits(0)._ColumnProps(14)=   "Column(2)._WidthInPix=6059"
      Splits(0)._ColumnProps(15)=   "Column(2)._EditAlways=0"
      Splits(0)._ColumnProps(16)=   "Column(2).Order=3"
      Splits(0)._ColumnProps(17)=   "Column(3).Width=185"
      Splits(0)._ColumnProps(18)=   "Column(3).DividerColor=0"
      Splits(0)._ColumnProps(19)=   "Column(3)._WidthInPix=106"
      Splits(0)._ColumnProps(20)=   "Column(3)._EditAlways=0"
      Splits(0)._ColumnProps(21)=   "Column(3)._ColStyle=1"
      Splits(0)._ColumnProps(22)=   "Column(3).Visible=0"
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
      _StyleDefs(40)  =   "Splits(0).Columns(1).Style:id=28,.parent=11"
      _StyleDefs(41)  =   "Splits(0).Columns(1).HeadingStyle:id=25,.parent=12"
      _StyleDefs(42)  =   "Splits(0).Columns(1).FooterStyle:id=26,.parent=13"
      _StyleDefs(43)  =   "Splits(0).Columns(1).EditorStyle:id=27,.parent=15"
      _StyleDefs(44)  =   "Splits(0).Columns(2).Style:id=62,.parent=11"
      _StyleDefs(45)  =   "Splits(0).Columns(2).HeadingStyle:id=59,.parent=12"
      _StyleDefs(46)  =   "Splits(0).Columns(2).FooterStyle:id=60,.parent=13"
      _StyleDefs(47)  =   "Splits(0).Columns(2).EditorStyle:id=61,.parent=15"
      _StyleDefs(48)  =   "Splits(0).Columns(3).Style:id=32,.parent=11,.alignment=2"
      _StyleDefs(49)  =   "Splits(0).Columns(3).HeadingStyle:id=29,.parent=12"
      _StyleDefs(50)  =   "Splits(0).Columns(3).FooterStyle:id=30,.parent=13"
      _StyleDefs(51)  =   "Splits(0).Columns(3).EditorStyle:id=31,.parent=15"
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
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Caption         =   "Resultados"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   30
      TabIndex        =   26
      Top             =   9075
      Width           =   13215
   End
   Begin VB.Label lbldeter 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Caption         =   "Marque el producto a suministrar y pulse + para añadirlo al pedido"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   90
      TabIndex        =   18
      Top             =   2730
      Width           =   13155
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0FF&
      Caption         =   "Productos a suministrar : 0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Index           =   0
      Left            =   60
      TabIndex        =   17
      Top             =   5325
      Width           =   13230
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "Nuevo Pedido de Producto Controlado"
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
      Top             =   30
      Width           =   4020
   End
   Begin VB.Label lblsubtitulo 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "Especifique el detalle del pedido"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   90
      TabIndex        =   15
      Top             =   270
      Width           =   2280
   End
   Begin VB.Shape fondo 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   525
      Left            =   0
      Top             =   0
      Width           =   13500
   End
End
Attribute VB_Name = "frmPC_Detalle"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public PK As Long
Dim x As New XArrayDB
Const filas As Integer = 30
Const Col As Integer = 3
Private Enum COLS
    PARAMETRO = 0
    REQUISITO = 1
    VALOR = 2
    unidades = 3
End Enum


Private Sub chkFechaCaducidadNA_Click()
    If chkFechaCaducidadNA.Value = Checked Then
        fCaducidad.Enabled = False
    Else
        fCaducidad.Enabled = True
    End If
End Sub

Private Sub cmbPedidos_change()
    If cmbPedidos.getTEXTO <> "" Then
        txtDatos(1) = cmbPedidos.getTEXTO
        Dim oCP As New clsClientes_pedidos
        oCP.Carga cmbPedidos.getPK_SALIDA
        If IsDate(oCP.getFECHA_PEDIDO) Then
            fechaCliente = oCP.getFECHA_PEDIDO
        End If
    End If
End Sub

Private Sub cmdFactura_Click()
   On Error GoTo cmdfactura_Click_Error

    If txtDOC_ID <> "" Then
        Dim oDP As New clsDocs_pago
        oDP.generar_factura txtDOC_ID, False, "", "rptFactura"
        Set oDP = Nothing
    End If

   On Error GoTo 0
   Exit Sub

cmdfactura_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure cmdFactura_Click of Formulario frmSuministros_Lote"
End Sub

Private Sub cmbCliente_Change()
    If cmbCliente.getTEXTO <> "" Then
        pedidos cmbCliente.getPK_SALIDA
    Else
        cmbCliente.Limpiar
    End If
End Sub

Private Sub cmbproducto_Change()
    If cmbProducto.getTEXTO <> "" Then
        ' Codigo del reactivo
        Dim oTb As New clsTipos_bote_ex
        oTb.cargar cmbProducto.getPK_SALIDA
        txtCodigo = oTb.getCODIGO_IDEN_LOTE
        calcular_lote
        cargar_reactivos_stock (cmbProducto.getPK_SALIDA)
        cargar_parametros (cmbProducto.getPK_SALIDA)
    End If
End Sub
Private Sub cmdcancel_Click()
    Unload Me
End Sub
Private Sub cmdElimina_Click()
    If listaPedido.ListItems.Count > 0 Then
        With listaReactivos.ListItems.Add(, , listaPedido.ListItems(listaPedido.selectedItem.Index).Text)
            .SubItems(1) = listaPedido.ListItems(listaPedido.selectedItem.Index).SubItems(1)
            .SubItems(2) = listaPedido.ListItems(listaPedido.selectedItem.Index).SubItems(2)
            .SubItems(3) = listaPedido.ListItems(listaPedido.selectedItem.Index).SubItems(3)
            .SubItems(4) = listaPedido.ListItems(listaPedido.selectedItem.Index).SubItems(4)
        End With
        listaReactivos.Sorted = True
        listaReactivos.SortKey = 1
        
        listaPedido.ListItems.Remove listaPedido.selectedItem.Index
        If listaPedido.ListItems.Count > 0 Then
            Set listaPedido.selectedItem = listaPedido.ListItems(listaPedido.selectedItem.Index)
            listaPedido.selectedItem.EnsureVisible
        End If
        
        Label2(0).Caption = "Productos a suministrar : " & listaPedido.ListItems.Count
        calcular_lote
        calcularIdentificacion
    End If
End Sub

Private Sub cmdInserta_Click()
    If listaReactivos.ListItems.Count > 0 Then
        ' Lote columna 4
        Dim i As Integer
        For i = 1 To listaPedido.ListItems.Count
            If listaReactivos.ListItems(listaReactivos.selectedItem.Index).SubItems(4) <> listaPedido.ListItems(i).SubItems(4) Then
                MsgBox "No se pueden suministrar reactivos de distintos lotes.", vbCritical, App.Title
                Exit Sub
            End If
        Next
        With listaPedido.ListItems.Add(, , listaReactivos.ListItems(listaReactivos.selectedItem.Index).Text)
            .SubItems(1) = listaReactivos.ListItems(listaReactivos.selectedItem.Index).SubItems(1)
            .SubItems(2) = listaReactivos.ListItems(listaReactivos.selectedItem.Index).SubItems(2)
            .SubItems(3) = listaReactivos.ListItems(listaReactivos.selectedItem.Index).SubItems(3)
            .SubItems(4) = listaReactivos.ListItems(listaReactivos.selectedItem.Index).SubItems(4)
        End With
        
        listaReactivos.ListItems.Remove listaReactivos.selectedItem.Index
        If listaReactivos.ListItems.Count > 0 Then
            Set listaReactivos.selectedItem = listaReactivos.ListItems(listaReactivos.selectedItem.Index)
            listaReactivos.selectedItem.EnsureVisible
        End If
        
        Label2(0).Caption = "Productos a suministrar : " & listaPedido.ListItems.Count
        calcular_lote
        calcularIdentificacion
    End If
End Sub

Private Sub cmdok_Click()
    On Error GoTo fallo
    If validar = True Then
      If PK <> 0 Then
        If MsgBox("Va a modificar el pedido. ¿Esta seguro?", vbQuestion + vbYesNo, App.Title) = vbNo Then
            Exit Sub
        End If
      End If
      Me.MousePointer = 11
      Dim oPC As New clsPc_pedido
      Dim i As Integer
      Dim LOTE As Long
      With oPC
        .setCLIENTE_ID = cmbCliente.getPK_SALIDA
        .setFECHA_PEDIDO = Format(fechaPedido, "yyyy-mm-dd")
        .setTIPO_BOTE_EX_ID = cmbProducto.getPK_SALIDA
        .setIDENTIFICACION = txtDatos(0)
        .setCANTIDAD = listaPedido.ListItems.Count
        .setPEDIDO_CLIENTE = txtDatos(1)
        .setFECHA_CLIENTE = Format(fechaCliente, "yyyy-mm-dd")
        .setPRECIO = moneda_bd(txtPrecio)
        If cmbPedidos.getTEXTO = "" Then
            .setPEDIDO_ID = 0
        Else
            .setPEDIDO_ID = cmbPedidos.getPK_SALIDA
        End If
        ' Nuevos campos
        If txtNumCertificado = "" Then
            .setNCERTIFICADO = 0
        Else
            .setNCERTIFICADO = txtNumCertificado
        End If
        If cmbNormativa.getTEXTO = "" Then
            .setNORMATIVA_ID = 0
        Else
            .setNORMATIVA_ID = cmbNormativa.getPK_SALIDA
        End If
        If chkFechaCaducidadNA.Value = Checked Then
            .setF_CADUCIDAD = "NULL"
        Else
            .setF_CADUCIDAD = "'" & Format(fCaducidad.Value, "yyyy-mm-dd") & "'"
        End If
        .setIDENTIFICACION2 = txtIdentificacion
        .setINFORMES = txtInformeFabricante
        If PK = 0 Then
            LOTE = .Insertar
        Else
            .Modificar (PK)
            LOTE = PK
        End If
        ' Reactivos
        If LOTE <> 0 Then
            Dim oPCR As New clsPc_rex
            oPCR.Eliminar LOTE
            For i = 1 To listaPedido.ListItems.Count
              With oPCR
                .setPEDIDO_ID = LOTE
                .setBOTE_EX_ID = listaPedido.ListItems(i).Text
                .setORDEN = i
                .Insertar fechaPedido
              End With
            Next
           Else
           Exit Sub
        End If
        ' Parametros
        insertar_parametros LOTE
      Me.MousePointer = 0
      End With
      If PK = 0 Then
          MsgBox "El pedido se ha introducido correctamente.", vbOKOnly + vbInformation, App.Title '      Else
      Else
          MsgBox "El pedido se ha modificado correctamente.", vbOKOnly + vbInformation, App.Title
      End If
      Unload Me
    End If
    Exit Sub
fallo:
      Me.MousePointer = 0
    error_grave (Err.Description)
End Sub

Private Sub Form_Load()
    log (Me.Name)
    cargar_botones Me
    Call cabecera
    Call cargar_combos
    inicializar_grid
    fechaPedido = Date
    fechaCliente = Date
    fCaducidad = Date + 365
    If PK <> 0 Then
        cargar_lote
    Else
        ' Calcular nuevo numero de certificado
        Dim oPC As New clsPc_pedido
        txtNumCertificado = oPC.maxNCERTIFICADO(Date)
        Set oPC = Nothing
    End If
End Sub
Private Sub calcularIdentificacion()
    Dim i As Integer
    Dim MINIMO As Long
    Dim MAXIMO As Long
    Dim aux As Long
   On Error GoTo calcularIdentificacion_Error

    aux = 0
    MINIMO = 0
    MAXIMO = 0
    listaPedido.Sorted = True
    listaPedido.SortKey = 0
    Dim correlativo As Boolean
    correlativo = True
            

    For i = 1 To listaPedido.ListItems.Count
        If i = 1 Then
            aux = CLng(listaPedido.ListItems(i).Text)
            MINIMO = aux
            MAXIMO = aux
        Else
            If (CLng(listaPedido.ListItems(i - 1).Text) + 1) <> CLng(listaPedido.ListItems(i).Text) Then
                correlativo = False
            End If
        End If
        If MINIMO > CLng(listaPedido.ListItems(i).Text) Then
            MINIMO = CLng(listaPedido.ListItems(i).Text)
        End If
        If MAXIMO < CLng(listaPedido.ListItems(i).Text) Then
            MAXIMO = CLng(listaPedido.ListItems(i).Text)
        End If
    Next
    Dim salida As String
    If correlativo Then
        If MINIMO = 0 And MAXIMO = 0 Then
            salida = "N/A."
        Else
            If MINIMO = MAXIMO Or MINIMO = 0 Or MAXIMO = 0 Then
                salida = MINIMO
            Else
                If MINIMO <> MAXIMO Then
                    salida = MINIMO & " - " & MAXIMO
                End If
            End If
        End If
    Else
        For i = 1 To listaPedido.ListItems.Count
            If salida <> "" Then
                salida = salida & ","
            End If
            salida = salida & listaPedido.ListItems(i).Text
        Next
    End If
    txtIdentificacion = salida
        

   On Error GoTo 0
   Exit Sub

calcularIdentificacion_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure calcularIdentificacion of Formulario frmPC_Detalle"
        
End Sub
Private Sub cabecera()
    With listaReactivos.ColumnHeaders
        .Add , , "Número", 800, lvwColumnLeft
        .Add , , "Reactivo", 5600, lvwColumnCenter
        .Add , , "F.Pedido", 1100, lvwColumnCenter
        .Add , , "F.Caducidad", 1100, lvwColumnCenter
        .Add , , "Lote", 2700, lvwColumnCenter
    End With
    With listaPedido.ColumnHeaders
        .Add , , "Número", 800, lvwColumnLeft
        .Add , , "Reactivo", 5600, lvwColumnCenter
        .Add , , "F.Pedido", 1100, lvwColumnCenter
        .Add , , "F.Caducidad", 1100, lvwColumnCenter
        .Add , , "Lote", 2700, lvwColumnCenter
    End With
End Sub

Private Sub listaPedido_DblClick()
    If listaPedido.ListItems.Count > 0 Then
        frmREX_Bote_Modificacion.PK = CLng(listaPedido.ListItems(listaPedido.selectedItem.Index).Text)
        frmREX_Bote_Modificacion.Show 1
        cargar_reactivos_stock cmbProducto.getPK_SALIDA
    End If
End Sub

Private Sub listaReactivos_DblClick()
    If listaReactivos.ListItems.Count > 0 Then
        frmREX_Bote_Modificacion.PK = CLng(listaReactivos.ListItems(listaReactivos.selectedItem.Index).Text)
        frmREX_Bote_Modificacion.Show 1
        cargar_reactivos_stock cmbProducto.getPK_SALIDA
    End If
End Sub

Private Sub txtdatos_GotFocus(Index As Integer)
    txtDatos(Index).BackColor = &H80C0FF
    txtDatos(Index).SelStart = 0
    txtDatos(Index).SelLength = Len(txtDatos(Index))
End Sub

Private Sub txtDatos_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys "{Tab}", True
        KeyAscii = 0 ' Para evitar el "bip" del sistema
    End If
End Sub
Private Sub txtdatos_LostFocus(Index As Integer)
    txtDatos(Index).BackColor = vbWhite
End Sub
Private Function validar() As Boolean
    validar = True
    If cmbProducto.getTEXTO = "" Then
        MsgBox "Debe seleccionar el producto a suministrar.", vbInformation, App.Title
        cmbProducto.SetFocus
        validar = False
        Exit Function
    End If
    If txtDatos(0) = "" Then
        MsgBox "Debe especificar el número de lote.", vbInformation, App.Title
        txtDatos(0).SetFocus
        validar = False
        Exit Function
    End If
    If cmbCliente.getTEXTO = "" Then
        MsgBox "Debe indicar algún cliente.", vbInformation, App.Title
        clientes.SetFocus
        validar = False
    End If
    If Trim(txtPrecio) = "" Then
        MsgBox "Debe introducir el precio del Bote.", vbInformation, App.Title
        txtPrecio.SetFocus
        validar = False
        Exit Function
    End If
    If listaPedido.ListItems.Count = 0 Then
        MsgBox "Debe añadir algún producto al pedido.", vbInformation, App.Title
        listaPedido.SetFocus
        validar = False
    End If
    If Trim(txtNumCertificado) <> "" Then
        If Not IsNumeric(txtNumCertificado) Then
            MsgBox "Debe introducir el Numero de Certificado correctamente o dejarlo en blanco.", vbInformation, App.Title
            txtNumCertificado.SetFocus
            validar = False
            Exit Function
        
        End If
    End If
End Function

Private Sub cargar_combos()
    cargar_tipos_reactivos
    llenar_combo cmbCliente, New clsCliente, 0, frmClientes, ""
    llenar_combo cmbNormativa, New clsCa_normas, 0, frmCA_Normas, ""
    pedidos 0
End Sub
Private Sub cargar_tipos_reactivos()
    cmbProducto.Limpiar
    Dim consulta As String
    consulta = " SELECT DISTINCT TB.ID_TIPO_BOTE_EX,T.NOMBRE " & _
               "   FROM TIPOS_REACTIVO_EX T, TIPOS_BOTE_EX TB " & _
               "  WHERE T.ID_TIPO_REACTIVO_EX = TB.TIPO_REACTIVO_EX_ID " & _
               "    AND TB.TIPO_M_REFERENCIA_ID = 7"
    Dim conn As ADODB.Connection
    If CrearConexionGlobal(conn, "", "") = True Then ' CONECTAR EL RS
        With cmbProducto
            .setCONN = conn
            .setFK_CAMPO = ""
            .setFK_VALOR = 0
            .setTABLA = "TIPOS_BOTE_EX"
            .setDESCRIPCION = "Producto a suministrar"
            .setPK = "TB.ID_TIPO_BOTE_EX"
            .setCAMPO = "T.NOMBRE"
            .setFILTRO = ""
            .setQUERY = consulta
            .setMUESTRA_DETALLE = True
            Set .FORMULARIO = frmREX_Bote
        End With
    End If
End Sub
Private Sub cargar_reactivos_stock(Reactivo As Long)
    Dim rs As ADODB.Recordset
    Dim consulta As String
    listaReactivos.ListItems.Clear
    consulta = "SELECT A.ID_BOTE_EX, " & _
               "       C.NOMBRE, " & _
               "       A.FECHA_PEDIDO, " & _
               "       A.FECHA_CADUCIDAD, " & _
               "       A.LOTE " & _
               " FROM BOTES_EX A, " & _
               "      TIPOS_BOTE_EX B, " & _
               "      TIPOS_REACTIVO_EX C " & _
               " WHERE A.TIPO_BOTE_EX_ID = B.ID_TIPO_BOTE_EX " & _
               "   AND B.TIPO_REACTIVO_EX_ID = C.ID_TIPO_REACTIVO_EX " & _
               "   AND A.TIPO_BOTE_EX_ID = " & Reactivo & _
               "   AND A.ANULADO = 0 AND A.FINALIZADO  = 0 " & _
               "   AND A.ABIERTO = 0 " & _
               " ORDER BY A.ID_BOTE_EX ASC"
    Set rs = datos_bd(consulta)
    If rs.RecordCount > 0 Then
        While Not rs.EOF
            With listaReactivos.ListItems.Add(, , Format(rs(0), "00000"))
                .SubItems(1) = rs(1)
                .SubItems(2) = Format(rs(2), "dd-mm-yyyy")
                .SubItems(3) = Format(rs(3), "dd-mm-yyyy")
                .SubItems(4) = rs(4)
            End With
            rs.MoveNext
        Wend
    End If
End Sub

Private Sub cargar_lote()
    ' Detalle del Pedido
    Dim oPC As New clsPc_pedido
    With oPC
        .Carga PK
        cmbProducto.MostrarElemento .getTIPO_BOTE_EX_ID
        txtDatos(0) = .getIDENTIFICACION
        If IsDate(.getFECHA_PEDIDO) Then
            fechaPedido = .getFECHA_PEDIDO
        End If
        cmbCliente.MostrarElemento .getCLIENTE_ID
        txtDatos(1) = .getPEDIDO_CLIENTE
        If IsDate(.getFECHA_CLIENTE) Then
            fechaCliente = .getFECHA_CLIENTE
        End If
        txtPrecio = moneda(.getPRECIO)
        cmbPedidos.MostrarElemento .getPEDIDO_ID
        txtDOC_ID = .getDOC_ID
        If .getDOC_ID > 1 Then
            cmdFactura.Visible = True
            cmdok.Visible = False
        End If
        ' Nuevos campos
        If .getNCERTIFICADO <> 0 Then
            txtNumCertificado = .getNCERTIFICADO
        End If
        If .getF_CADUCIDAD <> "" Then
            chkFechaCaducidadNA.Value = Unchecked
            fCaducidad = .getF_CADUCIDAD
        End If
        cmbNormativa.MostrarElemento .getNORMATIVA_ID
        txtIdentificacion = .getIDENTIFICACION2
        txtInformeFabricante = .getINFORMES
    End With
    Set oPC = Nothing
    ' Reactivos
    Dim oPCD As New clsPc_rex
    Dim rs As ADODB.Recordset
    Set rs = oPCD.Listado(PK)
    If rs.RecordCount > 0 Then
        Do
            With listaPedido.ListItems.Add(, , Format(rs(0), "00000"))
                .SubItems(1) = rs(1)
                .SubItems(2) = Format(rs(2), "dd-mm-yyyy")
                .SubItems(3) = Format(rs(3), "dd-mm-yyyy")
                .SubItems(4) = rs(4)
            End With
            rs.MoveNext
        Loop Until rs.EOF
    End If
    Label2(0).Caption = "Productos a suministrar : " & listaPedido.ListItems.Count
End Sub

Private Sub calcular_lote()
    txtDatos(0) = ""
    If listaPedido.ListItems.Count > 0 Then
        txtDatos(0) = txtCodigo & "-" & listaPedido.ListItems(1).SubItems(4) & "-" & Format(fechaPedido, "yy")
    End If
End Sub
Private Sub inicializar_grid()
    x.ReDim 0, filas, 0, Col
    x.Clear
    Set grid.Array = x
    grid.Refresh
End Sub
Private Sub cargar_parametros(ID As Long)
    x.Clear
        
    Dim rs As ADODB.Recordset
    Dim oTBP As New clsTipos_bote_ex_parametros
    Dim oPPC As New clsPc_parametros
    If PK = 0 Then
        Set rs = oTBP.Listado(ID)
    Else
        Set rs = oPPC.Listado(PK)
    End If
    If rs.RecordCount > 0 Then
    Dim i As Integer
    i = 0
        Do
            x(i, COLS.PARAMETRO) = CStr(rs(0))
            x(i, COLS.REQUISITO) = CStr(rs(1))
            x(i, COLS.VALOR) = CStr(rs(2))
            x(i, COLS.unidades) = CStr(rs(3))
            rs.MoveNext
            i = i + 1
        Loop Until rs.EOF
    End If
    Set oTBP = Nothing
    Set rs = Nothing
    
    grid.Refresh
End Sub

Private Sub insertar_parametros(pedido As Long)
    ' Evidencias
    Dim oTBP As New clsPc_parametros
   On Error GoTo insertar_parametros_Error

    oTBP.Eliminar PK
    Dim i As Integer
    For i = x.LowerBound(1) To x.UpperBound(1)
        If Trim(x.Value(i, COLS.PARAMETRO)) <> "" Then
            With oTBP
                .setPEDIDO_ID = pedido
                .setORDEN = i
                .setPARAMETRO = x.Value(i, COLS.PARAMETRO)
                .setREQUISITO = x.Value(i, COLS.REQUISITO)
                .setVALOR = x.Value(i, COLS.VALOR)
                .setUNIDADES = x.Value(i, COLS.unidades)
                .Insertar
            End With
        End If
    Next
    Set oTBP = Nothing

   On Error GoTo 0
   Exit Sub

insertar_parametros_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure insertar_parametros of Formulario frmREX_Bote"
End Sub

Private Sub pedidos(ID As Long)
    Dim filtro As String
    If ID <> 0 Then
        If listaReactivos.ListItems.Count > 0 Then
'            filtro = " AND CLIENTE_ID = " & ID & " AND FECHA_BAJA >= '" & Format(listaReactivos.ListItems(listaReactivos.selectedItem.Index).SubItems(4), "YYYY-MM-DD") & "'"
            filtro = " AND CLIENTE_ID = " & ID & " AND FECHA_BAJA >= '" & Format(fechaPedido, "YYYY-MM-DD") & "'"
            
        Else
            filtro = " AND CLIENTE_ID = " & ID
        End If
    End If
    cmbPedidos.Limpiar
    llenar_combo cmbPedidos, New clsClientes_pedidos, 0, frmClientes_Pedidos, filtro
End Sub

Private Sub txtprecio_LostFocus()
    If txtPrecio <> "" Then
        txtPrecio = moneda(txtPrecio)
    End If
End Sub
