VERSION 5.00
Object = "{A49CE0E0-C0F9-11D2-B0EA-00A024695830}#1.0#0"; "tidate8.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{CEA3B7F6-2847-4E5E-A551-DB7A62489D44}#46.0#0"; "miCombo.ocx"
Object = "{77DDF307-D82B-4757-8B3A-106EC9D75D4B}#8.0#0"; "tdbg8.ocx"
Begin VB.Form frmCE_Recepcion 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Recepción de Control de Eficacia"
   ClientHeight    =   11535
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   14595
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "frmCE_Recepcion_Nuevo2.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   11535
   ScaleWidth      =   14595
   Begin TrueDBGrid80.TDBDropDown tProducto 
      Height          =   2955
      Left            =   6300
      TabIndex        =   41
      Top             =   7380
      Width           =   7905
      _ExtentX        =   13944
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
   Begin TDBDate6Ctl.TDBDate TDBDate1 
      Height          =   285
      Left            =   11115
      TabIndex        =   40
      Top             =   10710
      Visible         =   0   'False
      Width           =   780
      _Version        =   65536
      _ExtentX        =   1376
      _ExtentY        =   503
      Calendar        =   "frmCE_Recepcion_Nuevo2.frx":2AFA
      Caption         =   "frmCE_Recepcion_Nuevo2.frx":2C12
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmCE_Recepcion_Nuevo2.frx":2C7E
      Keys            =   "frmCE_Recepcion_Nuevo2.frx":2C9C
      Spin            =   "frmCE_Recepcion_Nuevo2.frx":2CFA
      AlignHorizontal =   0
      AlignVertical   =   2
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
      Text            =   "27/10/2010"
      ValidateMode    =   0
      ValueVT         =   7
      Value           =   40478
      CenturyMode     =   0
   End
   Begin Geslab.ControlPanelXP cpDatos 
      Height          =   2865
      Left            =   45
      TabIndex        =   25
      Top             =   585
      Width           =   14505
      _ExtentX        =   25585
      _ExtentY        =   5054
      Caption         =   "Datos de recepción del control de eficacia"
      BackColor       =   12632256
      HeaderColor     =   8421504
      Object.Height          =   2865
      Begin MSDataListLib.DataCombo cmbCentro 
         Bindings        =   "frmCE_Recepcion_Nuevo2.frx":2D22
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
         TabIndex        =   2
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
         TabIndex        =   1
         Top             =   855
         Width           =   7890
         _ExtentX        =   13917
         _ExtentY        =   582
      End
      Begin pryCombo.miCombo cmbLote 
         Height          =   330
         Left            =   9225
         TabIndex        =   12
         Top             =   2430
         Width           =   5130
         _ExtentX        =   9049
         _ExtentY        =   582
      End
      Begin MSDataListLib.DataCombo cmbenvases 
         Height          =   315
         Left            =   9225
         TabIndex        =   10
         Top             =   2070
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
      Begin MSComCtl2.DTPicker fprocesado 
         Height          =   330
         Left            =   2070
         TabIndex        =   8
         Top             =   2070
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
         Format          =   16515073
         CurrentDate     =   38000
         MinDate         =   2
      End
      Begin MSDataListLib.DataCombo cmbrealizada 
         Height          =   315
         Left            =   9225
         TabIndex        =   7
         Top             =   1710
         Width           =   4095
         _ExtentX        =   7223
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
         Left            =   2070
         TabIndex        =   6
         Top             =   1710
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
      Begin MSComCtl2.DTPicker fecha 
         Height          =   330
         Left            =   12060
         TabIndex        =   3
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
         Format          =   16515073
         CurrentDate     =   38000
         MinDate         =   2
      End
      Begin VB.TextBox txtDatos 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00BCF3EF&
         Height          =   330
         Index           =   0
         Left            =   12060
         MaxLength       =   255
         TabIndex        =   4
         Top             =   855
         Width           =   2280
      End
      Begin VB.CheckBox chkSinEspecificar 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Caption         =   "Sin Especificar"
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   3465
         TabIndex        =   9
         Top             =   2070
         Width           =   1365
      End
      Begin VB.TextBox txtDatos 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Enabled         =   0   'False
         ForeColor       =   &H000000FF&
         Height          =   330
         Index           =   1
         Left            =   2070
         TabIndex        =   11
         Text            =   "Realizar análisis"
         Top             =   2430
         Width           =   3810
      End
      Begin pryCombo.miCombo cmbClientes 
         Height          =   330
         Left            =   1170
         TabIndex        =   0
         Top             =   495
         Width           =   7890
         _ExtentX        =   13917
         _ExtentY        =   582
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Centro"
         Height          =   195
         Index           =   22
         Left            =   10935
         TabIndex        =   43
         Top             =   1260
         Width           =   465
      End
      Begin VB.Image imgPedidos 
         Height          =   300
         Left            =   8430
         Picture         =   "frmCE_Recepcion_Nuevo2.frx":2D68
         Stretch         =   -1  'True
         Top             =   1230
         Width           =   315
      End
      Begin VB.Image Image1 
         Height          =   300
         Left            =   8100
         Picture         =   "frmCE_Recepcion_Nuevo2.frx":3632
         Stretch         =   -1  'True
         Top             =   1215
         Width           =   255
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Pedido"
         Height          =   195
         Index           =   0
         Left            =   225
         TabIndex        =   36
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
         TabIndex        =   35
         Top             =   540
         Width           =   945
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Baño"
         Height          =   195
         Index           =   1
         Left            =   225
         TabIndex        =   34
         Top             =   900
         Width           =   375
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Cliente"
         Height          =   195
         Index           =   3
         Left            =   225
         TabIndex        =   33
         Top             =   540
         Width           =   480
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Orden Compra"
         Height          =   195
         Index           =   9
         Left            =   10935
         TabIndex        =   32
         Top             =   900
         Width           =   1035
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Realizada por"
         Height          =   195
         Index           =   7
         Left            =   8100
         TabIndex        =   31
         Top             =   1755
         Width           =   975
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Entregada por"
         Height          =   195
         Index           =   4
         Left            =   225
         TabIndex        =   30
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
         Caption         =   "Procesado de las piezas"
         Height          =   195
         Index           =   16
         Left            =   225
         TabIndex        =   29
         Top             =   2115
         Width           =   1755
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Envase"
         Height          =   195
         Index           =   5
         Left            =   8100
         TabIndex        =   28
         Top             =   2130
         Width           =   540
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Lote Probetas"
         Height          =   195
         Index           =   18
         Left            =   8100
         TabIndex        =   27
         Top             =   2475
         Width           =   990
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Datos de Espesor"
         Height          =   195
         Index           =   17
         Left            =   225
         TabIndex        =   26
         Top             =   2475
         Width           =   1260
      End
   End
   Begin Geslab.ControlPanelXP cpprobetas 
      Height          =   7035
      Left            =   45
      TabIndex        =   37
      Top             =   3465
      Width           =   14505
      _ExtentX        =   23760
      _ExtentY        =   12409
      Caption         =   "Probetas y Ensayos"
      BackColor       =   12632256
      HeaderColor     =   8421504
      Object.Height          =   7035
      Begin TrueDBGrid80.TDBDropDown tProbetas 
         Height          =   2730
         Left            =   3915
         TabIndex        =   39
         Top             =   4230
         Width           =   8625
         _ExtentX        =   15214
         _ExtentY        =   4815
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
         Height          =   2730
         Left            =   90
         TabIndex        =   38
         Top             =   3915
         Width           =   14250
         _ExtentX        =   25135
         _ExtentY        =   4815
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
         Splits(0)._ColumnProps(1)=   "Column(0).Width=9208"
         Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
         Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=9128"
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
      Begin TrueDBGrid80.TDBGrid gridA 
         Height          =   3600
         Left            =   45
         TabIndex        =   14
         Top             =   3375
         Width           =   14415
         _ExtentX        =   25426
         _ExtentY        =   6350
         _LayoutType     =   4
         _RowHeight      =   16
         _WasPersistedAsPixels=   0
         Columns(0)._VlistStyle=   0
         Columns(0)._MaxComboItems=   5
         Columns(0).Caption=   "Tipo de Ensayo"
         Columns(0).DataField=   ""
         Columns(0).DropDown=   "tAnalisis"
         Columns(0).DropDown.vt=   8
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
         Columns(3)._VlistStyle=   4
         Columns(3)._MaxComboItems=   5
         Columns(3).Caption=   "Enac"
         Columns(3).DataField=   ""
         Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(4)._VlistStyle=   0
         Columns(4)._MaxComboItems=   5
         Columns(4).Caption=   "Des.Producto"
         Columns(4).DataField=   ""
         Columns(4).DropDown=   "tProducto"
         Columns(4).DropDown.vt=   8
         Columns(4)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(5)._VlistStyle=   0
         Columns(5)._MaxComboItems=   5
         Columns(5).Caption=   "F.Entrega"
         Columns(5).DataField=   ""
         Columns(5).NumberFormat=   "Short Date"
         Columns(5).ExternalEditor=   "TDBDate1"
         Columns(5).ExternalEditor.vt=   8
         Columns(5)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(6)._VlistStyle=   0
         Columns(6)._MaxComboItems=   5
         Columns(6).Caption=   "F.Procesado"
         Columns(6).DataField=   ""
         Columns(6).NumberFormat=   "Short Date"
         Columns(6).ExternalEditor=   "TDBDate1"
         Columns(6).ExternalEditor.vt=   8
         Columns(6)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(7)._VlistStyle=   0
         Columns(7)._MaxComboItems=   5
         Columns(7).Caption=   "Ref. Cliente"
         Columns(7).DataField=   ""
         Columns(7)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(8)._VlistStyle=   0
         Columns(8)._MaxComboItems=   5
         Columns(8).Caption=   "ID_ENSAYO"
         Columns(8).DataField=   ""
         Columns(8).NumberFormat=   "General Number"
         Columns(8).DropDown=   "tEstados"
         Columns(8).DropDown.vt=   8
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
         Splits(0)._ColumnProps(1)=   "Column(0).Width=6826"
         Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
         Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=6747"
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
         Splits(0)._ColumnProps(15)=   "Column(1).Visible=0"
         Splits(0)._ColumnProps(16)=   "Column(1).Order=2"
         Splits(0)._ColumnProps(17)=   "Column(2).Width=3254"
         Splits(0)._ColumnProps(18)=   "Column(2).DividerColor=0"
         Splits(0)._ColumnProps(19)=   "Column(2)._WidthInPix=3175"
         Splits(0)._ColumnProps(20)=   "Column(2)._EditAlways=0"
         Splits(0)._ColumnProps(21)=   "Column(2)._ColStyle=1"
         Splits(0)._ColumnProps(22)=   "Column(2).Order=3"
         Splits(0)._ColumnProps(23)=   "Column(2).AutoDropDown=1"
         Splits(0)._ColumnProps(24)=   "Column(2).DropDownList=1"
         Splits(0)._ColumnProps(25)=   "Column(2).AutoCompletion=1"
         Splits(0)._ColumnProps(26)=   "Column(3).Width=847"
         Splits(0)._ColumnProps(27)=   "Column(3).DividerColor=0"
         Splits(0)._ColumnProps(28)=   "Column(3)._WidthInPix=767"
         Splits(0)._ColumnProps(29)=   "Column(3)._EditAlways=0"
         Splits(0)._ColumnProps(30)=   "Column(3).AllowSizing=0"
         Splits(0)._ColumnProps(31)=   "Column(3)._ColStyle=1"
         Splits(0)._ColumnProps(32)=   "Column(3).AllowFocus=0"
         Splits(0)._ColumnProps(33)=   "Column(3).Order=4"
         Splits(0)._ColumnProps(34)=   "Column(4).Width=2937"
         Splits(0)._ColumnProps(35)=   "Column(4).DividerColor=0"
         Splits(0)._ColumnProps(36)=   "Column(4)._WidthInPix=2858"
         Splits(0)._ColumnProps(37)=   "Column(4)._EditAlways=0"
         Splits(0)._ColumnProps(38)=   "Column(4)._ColStyle=0"
         Splits(0)._ColumnProps(39)=   "Column(4).Order=5"
         Splits(0)._ColumnProps(40)=   "Column(4).AutoDropDown=1"
         Splits(0)._ColumnProps(41)=   "Column(4).DropDownList=1"
         Splits(0)._ColumnProps(42)=   "Column(4).AutoCompletion=1"
         Splits(0)._ColumnProps(43)=   "Column(5).Width=2196"
         Splits(0)._ColumnProps(44)=   "Column(5).DividerColor=0"
         Splits(0)._ColumnProps(45)=   "Column(5)._WidthInPix=2117"
         Splits(0)._ColumnProps(46)=   "Column(5)._EditAlways=0"
         Splits(0)._ColumnProps(47)=   "Column(5)._ColStyle=1"
         Splits(0)._ColumnProps(48)=   "Column(5).Order=6"
         Splits(0)._ColumnProps(49)=   "Column(6).Width=2117"
         Splits(0)._ColumnProps(50)=   "Column(6).DividerColor=0"
         Splits(0)._ColumnProps(51)=   "Column(6)._WidthInPix=2037"
         Splits(0)._ColumnProps(52)=   "Column(6)._EditAlways=0"
         Splits(0)._ColumnProps(53)=   "Column(6).AllowSizing=0"
         Splits(0)._ColumnProps(54)=   "Column(6)._ColStyle=1"
         Splits(0)._ColumnProps(55)=   "Column(6).Order=7"
         Splits(0)._ColumnProps(56)=   "Column(6).AutoDropDown=1"
         Splits(0)._ColumnProps(57)=   "Column(6).DropDownList=1"
         Splits(0)._ColumnProps(58)=   "Column(7).Width=2725"
         Splits(0)._ColumnProps(59)=   "Column(7).DividerColor=0"
         Splits(0)._ColumnProps(60)=   "Column(7)._WidthInPix=2646"
         Splits(0)._ColumnProps(61)=   "Column(7)._EditAlways=0"
         Splits(0)._ColumnProps(62)=   "Column(7).Order=8"
         Splits(0)._ColumnProps(63)=   "Column(8).Width=4260"
         Splits(0)._ColumnProps(64)=   "Column(8).DividerColor=0"
         Splits(0)._ColumnProps(65)=   "Column(8)._WidthInPix=4180"
         Splits(0)._ColumnProps(66)=   "Column(8)._EditAlways=0"
         Splits(0)._ColumnProps(67)=   "Column(8)._ColStyle=1"
         Splits(0)._ColumnProps(68)=   "Column(8).Visible=0"
         Splits(0)._ColumnProps(69)=   "Column(8).Order=9"
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
         Caption         =   "II. Ensayos (Pulse Botón derecho para verlos en detalle)"
         TabAction       =   2
         WrapCellPointer =   -1  'True
         MultipleLines   =   0
         EmptyRows       =   -1  'True
         CellTipsWidth   =   0
         MultiSelect     =   2
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
         _StyleDefs(54)  =   "Splits(0).Columns(3).Style:id=66,.parent=11,.alignment=2"
         _StyleDefs(55)  =   "Splits(0).Columns(3).HeadingStyle:id=63,.parent=12"
         _StyleDefs(56)  =   "Splits(0).Columns(3).FooterStyle:id=64,.parent=13"
         _StyleDefs(57)  =   "Splits(0).Columns(3).EditorStyle:id=65,.parent=15"
         _StyleDefs(58)  =   "Splits(0).Columns(4).Style:id=62,.parent=11,.alignment=0,.wraptext=0"
         _StyleDefs(59)  =   "Splits(0).Columns(4).HeadingStyle:id=59,.parent=12"
         _StyleDefs(60)  =   "Splits(0).Columns(4).FooterStyle:id=60,.parent=13"
         _StyleDefs(61)  =   "Splits(0).Columns(4).EditorStyle:id=61,.parent=15"
         _StyleDefs(62)  =   "Splits(0).Columns(5).Style:id=70,.parent=11,.alignment=2"
         _StyleDefs(63)  =   "Splits(0).Columns(5).HeadingStyle:id=67,.parent=12"
         _StyleDefs(64)  =   "Splits(0).Columns(5).FooterStyle:id=68,.parent=13"
         _StyleDefs(65)  =   "Splits(0).Columns(5).EditorStyle:id=69,.parent=15"
         _StyleDefs(66)  =   "Splits(0).Columns(6).Style:id=58,.parent=11,.alignment=2"
         _StyleDefs(67)  =   "Splits(0).Columns(6).HeadingStyle:id=55,.parent=12"
         _StyleDefs(68)  =   "Splits(0).Columns(6).FooterStyle:id=56,.parent=13"
         _StyleDefs(69)  =   "Splits(0).Columns(6).EditorStyle:id=57,.parent=15"
         _StyleDefs(70)  =   "Splits(0).Columns(7).Style:id=28,.parent=11,.bgcolor=&HC1FFFF&"
         _StyleDefs(71)  =   "Splits(0).Columns(7).HeadingStyle:id=25,.parent=12"
         _StyleDefs(72)  =   "Splits(0).Columns(7).FooterStyle:id=26,.parent=13"
         _StyleDefs(73)  =   "Splits(0).Columns(7).EditorStyle:id=27,.parent=15"
         _StyleDefs(74)  =   "Splits(0).Columns(8).Style:id=54,.parent=11,.alignment=2"
         _StyleDefs(75)  =   "Splits(0).Columns(8).HeadingStyle:id=51,.parent=12"
         _StyleDefs(76)  =   "Splits(0).Columns(8).FooterStyle:id=52,.parent=13"
         _StyleDefs(77)  =   "Splits(0).Columns(8).EditorStyle:id=53,.parent=15"
         _StyleDefs(78)  =   "Named:id=37:Normal"
         _StyleDefs(79)  =   ":id=37,.parent=0,.alignment=3,.bgcolor=&H80000018&,.bgpicMode=2,.bold=0"
         _StyleDefs(80)  =   ":id=37,.fontsize=825,.italic=0,.underline=0,.strikethrough=0,.charset=0"
         _StyleDefs(81)  =   ":id=37,.fontname=MS Sans Serif"
         _StyleDefs(82)  =   "Named:id=38:Heading"
         _StyleDefs(83)  =   ":id=38,.parent=37,.valignment=2,.bgcolor=&H80000016&,.fgcolor=&H80000012&"
         _StyleDefs(84)  =   ":id=38,.wraptext=-1,.bold=0,.fontsize=825,.italic=0,.underline=0"
         _StyleDefs(85)  =   ":id=38,.strikethrough=0,.charset=0"
         _StyleDefs(86)  =   ":id=38,.fontname=MS Sans Serif"
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
         _StyleDefs(98)  =   ":id=43,.parent=37,.bgcolor=&HFFFFFF&,.wraptext=-1"
         _StyleDefs(99)  =   "Named:id=44:OddRow"
         _StyleDefs(100) =   ":id=44,.parent=37,.bgcolor=&HD5ECF9&"
         _StyleDefs(101) =   "Named:id=47:RecordSelector"
         _StyleDefs(102) =   ":id=47,.parent=38"
         _StyleDefs(103) =   "Named:id=50:FilterBar"
         _StyleDefs(104) =   ":id=50,.parent=37"
      End
      Begin TrueDBGrid80.TDBGrid gridP 
         Height          =   2880
         Left            =   45
         TabIndex        =   13
         Top             =   450
         Width           =   14415
         _ExtentX        =   25426
         _ExtentY        =   5080
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
         Columns(2).Caption=   "Dimensión"
         Columns(2).DataField=   ""
         Columns(2).DropDown=   "tResponsables"
         Columns(2).DropDown.vt=   8
         Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(3)._VlistStyle=   0
         Columns(3)._MaxComboItems=   5
         Columns(3).Caption=   "Nº Probetas"
         Columns(3).DataField=   ""
         Columns(3).NumberFormat=   "General Number"
         Columns(3).ExternalEditor=   "TDBDate1"
         Columns(3).ExternalEditor.vt=   8
         Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(4)._VlistStyle=   0
         Columns(4)._MaxComboItems=   5
         Columns(4).Caption=   "Áreas"
         Columns(4).DataField=   ""
         Columns(4).NumberFormat=   "General Number"
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
         Columns(9)._VlistStyle=   0
         Columns(9)._MaxComboItems=   5
         Columns(9).Caption=   "DES_PRODUCTO"
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
         Splits(0).AllowRowSelect=   0   'False
         Splits(0).DividerColor=   12632256
         Splits(0).SpringMode=   0   'False
         Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
         Splits(0)._ColumnProps(0)=   "Columns.Count=10"
         Splits(0)._ColumnProps(1)=   "Column(0).Width=5503"
         Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
         Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=5424"
         Splits(0)._ColumnProps(4)=   "Column(0)._EditAlways=0"
         Splits(0)._ColumnProps(5)=   "Column(0)._ColStyle=1"
         Splits(0)._ColumnProps(6)=   "Column(0).Order=1"
         Splits(0)._ColumnProps(7)=   "Column(1).Width=7699"
         Splits(0)._ColumnProps(8)=   "Column(1).DividerColor=0"
         Splits(0)._ColumnProps(9)=   "Column(1)._WidthInPix=7620"
         Splits(0)._ColumnProps(10)=   "Column(1)._EditAlways=0"
         Splits(0)._ColumnProps(11)=   "Column(1)._ColStyle=1"
         Splits(0)._ColumnProps(12)=   "Column(1).Order=2"
         Splits(0)._ColumnProps(13)=   "Column(2).Width=4604"
         Splits(0)._ColumnProps(14)=   "Column(2).DividerColor=0"
         Splits(0)._ColumnProps(15)=   "Column(2)._WidthInPix=4524"
         Splits(0)._ColumnProps(16)=   "Column(2)._EditAlways=0"
         Splits(0)._ColumnProps(17)=   "Column(2)._ColStyle=1"
         Splits(0)._ColumnProps(18)=   "Column(2).Order=3"
         Splits(0)._ColumnProps(19)=   "Column(2).DropDownList=1"
         Splits(0)._ColumnProps(20)=   "Column(3).Width=3413"
         Splits(0)._ColumnProps(21)=   "Column(3).DividerColor=0"
         Splits(0)._ColumnProps(22)=   "Column(3)._WidthInPix=3334"
         Splits(0)._ColumnProps(23)=   "Column(3)._EditAlways=0"
         Splits(0)._ColumnProps(24)=   "Column(3)._ColStyle=8193"
         Splits(0)._ColumnProps(25)=   "Column(3).Order=4"
         Splits(0)._ColumnProps(26)=   "Column(4).Width=1958"
         Splits(0)._ColumnProps(27)=   "Column(4).DividerColor=0"
         Splits(0)._ColumnProps(28)=   "Column(4)._WidthInPix=1879"
         Splits(0)._ColumnProps(29)=   "Column(4)._EditAlways=0"
         Splits(0)._ColumnProps(30)=   "Column(4)._ColStyle=1"
         Splits(0)._ColumnProps(31)=   "Column(4).Order=5"
         Splits(0)._ColumnProps(32)=   "Column(5).Width=344"
         Splits(0)._ColumnProps(33)=   "Column(5).DividerColor=0"
         Splits(0)._ColumnProps(34)=   "Column(5)._WidthInPix=265"
         Splits(0)._ColumnProps(35)=   "Column(5)._EditAlways=0"
         Splits(0)._ColumnProps(36)=   "Column(5)._ColStyle=1"
         Splits(0)._ColumnProps(37)=   "Column(5).Order=6"
         Splits(0)._ColumnProps(38)=   "Column(6).Width=1640"
         Splits(0)._ColumnProps(39)=   "Column(6).DividerColor=0"
         Splits(0)._ColumnProps(40)=   "Column(6)._WidthInPix=1561"
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
         Splits(0)._ColumnProps(57)=   "Column(9).Width=2725"
         Splits(0)._ColumnProps(58)=   "Column(9).DividerColor=0"
         Splits(0)._ColumnProps(59)=   "Column(9)._WidthInPix=2646"
         Splits(0)._ColumnProps(60)=   "Column(9)._EditAlways=0"
         Splits(0)._ColumnProps(61)=   "Column(9).Visible=0"
         Splits(0)._ColumnProps(62)=   "Column(9).Order=10"
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
         MultiSelect     =   2
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
         _StyleDefs(52)  =   "Splits(0).Columns(3).Style:id=28,.parent=11,.alignment=2,.locked=-1"
         _StyleDefs(53)  =   "Splits(0).Columns(3).HeadingStyle:id=25,.parent=12"
         _StyleDefs(54)  =   "Splits(0).Columns(3).FooterStyle:id=26,.parent=13"
         _StyleDefs(55)  =   "Splits(0).Columns(3).EditorStyle:id=27,.parent=15"
         _StyleDefs(56)  =   "Splits(0).Columns(4).Style:id=58,.parent=11,.alignment=2"
         _StyleDefs(57)  =   "Splits(0).Columns(4).HeadingStyle:id=55,.parent=12"
         _StyleDefs(58)  =   "Splits(0).Columns(4).FooterStyle:id=56,.parent=13"
         _StyleDefs(59)  =   "Splits(0).Columns(4).EditorStyle:id=57,.parent=15"
         _StyleDefs(60)  =   "Splits(0).Columns(5).Style:id=62,.parent=11,.alignment=2,.bgcolor=&HC1FFFF&"
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
         _StyleDefs(76)  =   "Splits(0).Columns(9).Style:id=54,.parent=11"
         _StyleDefs(77)  =   "Splits(0).Columns(9).HeadingStyle:id=51,.parent=12"
         _StyleDefs(78)  =   "Splits(0).Columns(9).FooterStyle:id=52,.parent=13"
         _StyleDefs(79)  =   "Splits(0).Columns(9).EditorStyle:id=53,.parent=15"
         _StyleDefs(80)  =   "Named:id=37:Normal"
         _StyleDefs(81)  =   ":id=37,.parent=0,.alignment=3,.bgcolor=&H80000018&,.bgpicMode=2,.bold=0"
         _StyleDefs(82)  =   ":id=37,.fontsize=825,.italic=0,.underline=0,.strikethrough=0,.charset=0"
         _StyleDefs(83)  =   ":id=37,.fontname=MS Sans Serif"
         _StyleDefs(84)  =   "Named:id=38:Heading"
         _StyleDefs(85)  =   ":id=38,.parent=37,.valignment=2,.bgcolor=&H80000016&,.fgcolor=&H80000012&"
         _StyleDefs(86)  =   ":id=38,.wraptext=-1,.bold=0,.fontsize=825,.italic=0,.underline=0"
         _StyleDefs(87)  =   ":id=38,.strikethrough=0,.charset=0"
         _StyleDefs(88)  =   ":id=38,.fontname=MS Sans Serif"
         _StyleDefs(89)  =   "Named:id=39:Footing"
         _StyleDefs(90)  =   ":id=39,.parent=37,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
         _StyleDefs(91)  =   "Named:id=40:Selected"
         _StyleDefs(92)  =   ":id=40,.parent=37,.bgcolor=&H8080FF&,.fgcolor=&H0&,.bold=0,.fontsize=825"
         _StyleDefs(93)  =   ":id=40,.italic=0,.underline=0,.strikethrough=0,.charset=0"
         _StyleDefs(94)  =   ":id=40,.fontname=MS Sans Serif"
         _StyleDefs(95)  =   "Named:id=41:Caption"
         _StyleDefs(96)  =   ":id=41,.parent=38,.alignment=2"
         _StyleDefs(97)  =   "Named:id=42:HighlightRow"
         _StyleDefs(98)  =   ":id=42,.parent=37,.bgcolor=&H80000008&,.fgcolor=&H80000005&"
         _StyleDefs(99)  =   "Named:id=43:EvenRow"
         _StyleDefs(100) =   ":id=43,.parent=37,.bgcolor=&HFFFFFF&,.wraptext=-1"
         _StyleDefs(101) =   "Named:id=44:OddRow"
         _StyleDefs(102) =   ":id=44,.parent=37,.bgcolor=&HD5ECF9&"
         _StyleDefs(103) =   "Named:id=47:RecordSelector"
         _StyleDefs(104) =   ":id=47,.parent=38"
         _StyleDefs(105) =   "Named:id=50:FilterBar"
         _StyleDefs(106) =   ":id=50,.parent=37"
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
      Left            =   5175
      Locked          =   -1  'True
      MaxLength       =   255
      TabIndex        =   24
      Top             =   315
      Width           =   7140
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
      TabIndex        =   23
      Top             =   45
      Width           =   7140
   End
   Begin VB.TextBox txtob 
      Appearance      =   0  'Flat
      BackColor       =   &H00D1CAF4&
      ForeColor       =   &H00000000&
      Height          =   705
      Left            =   4365
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   21
      Top             =   10530
      Width           =   6630
   End
   Begin VB.CheckBox chkFechaCompleta 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   "Usar fecha de procesado completa en la referencia del cliente"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   3105
      TabIndex        =   20
      Top             =   11250
      Width           =   5325
   End
   Begin VB.CommandButton cmdborrar 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Borrar Probeta"
      Height          =   930
      Left            =   90
      Picture         =   "frmCE_Recepcion_Nuevo2.frx":3EFC
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   10575
      Width           =   1245
   End
   Begin VB.CommandButton cmdborrarensayo 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Borrar Ensayo"
      Height          =   930
      Left            =   1395
      Picture         =   "frmCE_Recepcion_Nuevo2.frx":47C6
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   10575
      Width           =   1245
   End
   Begin VB.CommandButton cmdok 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Aceptar"
      Height          =   960
      Left            =   12390
      Style           =   1  'Graphical
      TabIndex        =   17
      Top             =   10545
      Width           =   1050
   End
   Begin VB.CommandButton cmdSalir 
      BackColor       =   &H00E0E0E0&
      Caption         =   "ESC-Salir"
      Height          =   960
      Left            =   13470
      Style           =   1  'Graphical
      TabIndex        =   18
      Top             =   10545
      Width           =   1050
   End
   Begin TDBDate6Ctl.TDBDate TDBDate2 
      Height          =   285
      Left            =   10755
      TabIndex        =   42
      Top             =   11205
      Visible         =   0   'False
      Width           =   780
      _Version        =   65536
      _ExtentX        =   1376
      _ExtentY        =   503
      Calendar        =   "frmCE_Recepcion_Nuevo2.frx":5090
      Caption         =   "frmCE_Recepcion_Nuevo2.frx":51A8
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmCE_Recepcion_Nuevo2.frx":5214
      Keys            =   "frmCE_Recepcion_Nuevo2.frx":5232
      Spin            =   "frmCE_Recepcion_Nuevo2.frx":5290
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
      Text            =   "27/10/2010"
      ValidateMode    =   0
      ValueVT         =   7
      Value           =   40478
      CenturyMode     =   0
   End
   Begin VB.Label Label2 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Observaciones"
      ForeColor       =   &H00000000&
      Height          =   285
      Index           =   13
      Left            =   3150
      TabIndex        =   22
      Top             =   10755
      Width           =   1095
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
      TabIndex        =   19
      Top             =   120
      Width           =   3495
   End
   Begin VB.Shape fondo 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   1  'Opaque
      Height          =   600
      Left            =   0
      Top             =   -45
      Width           =   14535
   End
End
Attribute VB_Name = "frmCE_Recepcion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim xP As New XArrayDB
Dim xA As New XArrayDB
Dim xAnalisis As New XArrayDB
Dim xProbetas As New XArrayDB
Dim xProducto As New XArrayDB
Const filasP As Integer = 1000
Const ColP As Integer = 10
Private Enum ColsP
    DESIGNACION = 0
    MATERIAL = 1
    DIMENSION = 2
    nProbetas = 3
    AREAS = 4
    RECIBIDAS = 5
    IDENTIFICADA = 6
    IDEN_CLIENTE = 7
    IDEN_CANAGROSA = 8
    DES_PRODUCTO = 9
End Enum
Const filasA As Integer = 100
'M1356-I
'Const ColA As Integer = 8
Const ColA As Integer = 9
'M1356-F
Private Enum ColsA
    tipo_ensayo = 0
    NORMA = 1
    DESIGNACION = 2
    ENAC = 3
    producto = 4
'M1356-I
'    FPROCESO = 5
'    REFCLIENTE = 6
'    ID_TIPO_ENSAYO = 7
    fentrega = 5
    FPROCESO = 6
    REFCLIENTE = 7
    ID_TIPO_ENSAYO = 8
'M1356-F
End Enum

Private Sub chkFechaCompleta_Click()
    calcular_referencias
End Sub

Private Sub cmdBorrar_Click()
    On Error Resume Next
    Dim i As Integer
    Dim j As Integer
    Dim k As Integer
    Dim l As Integer
  For l = 0 To gridP.SelBookmarks.Count
    gridP.Bookmark = gridP.SelBookmarks(l)
    For i = 0 To ColP
'        gridP.SelBookmarks.Add gridP.Bookmark
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
'        gridP.SelBookmarks.Remove 0
    Next i
  Next l
    gridP.Refresh
    gridP.SetFocus
    gridA.Refresh
    calcular_referencias
End Sub

Private Sub cmdborrarensayo_Click()
    On Error Resume Next
    Dim i As Integer
    Dim j As Integer
    For j = 0 To gridA.SelBookmarks.Count
        gridA.Bookmark = gridA.SelBookmarks(j)
        For i = 0 To ColA
'            gridA.SelBookmarks.Add gridA.Bookmark
            xA(gridA.Bookmark, i) = ""
'            gridA.SelBookmarks.Remove 0
        Next i
    Next j
    gridA.Refresh
    gridA.SetFocus
    calcular_referencias
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

Private Sub chkSinEspecificar_Click()
    fprocesado.Value = Date
    If chkSinEspecificar.Value = Checked Then
'FP        fprocesado.value = "01/01/1900"
        fprocesado.Enabled = False
    Else
'FP        fprocesado.value = Date
        fprocesado.Enabled = True
    End If
    informar_fechas_procesado
    calcular_referencias
End Sub
Private Sub cmbClientes_change()
    cargar_banos
    cmdLimpiar_Click
End Sub

Private Sub cmdLimpiar_Click()
    cmbPedido.Text = ""
    cmbPedido.BoundText = ""
End Sub

Private Sub cmdok_Click()
   On Error GoTo cmdok_Click_Error
    gridP.Col = 0
    gridP.Row = 0
    gridA.Col = 0
    gridA.Row = 0

    If validar = True Then
        Me.MousePointer = 11
        Dim oce_recepcion As New clsCe_recepcion
        Dim RECEPCION As Long
        Dim i As Integer
        oce_recepcion.Calcular_Numero_Recepcion
        ' Generamos el registro de las muestras
        Dim oMuestra As New clsMuestra
        Dim oce_tipo_ensayo As New clsCe_tipos_ensayos
        Dim oTipo_analisis As New clsTipos_analisis
        Dim oDatos_especificos As New clsDatos_valores
        Dim oTDA As New clsTipos_datos_analisis
        Dim oBANO As New clsBanos
        Dim oCe_resultados As New clsCe_resultados
        Dim muestra As Long
        Dim rs As ADODB.Recordset
        Dim indice As Integer
        Dim oCE_EQUIPO As New clsCe_recepcion_equipos
        
        For i = 0 To filasA
         If Not IsEmpty(xA(i, ColsA.ID_TIPO_ENSAYO)) Then
          If Trim(xA(i, ColsA.ID_TIPO_ENSAYO)) <> "" Then
            oce_tipo_ensayo.Carga (CLng(xA(i, ColsA.ID_TIPO_ENSAYO)))
            oTipo_analisis.CARGAR (oce_tipo_ensayo.getTIPO_ANALISIS_ID)
            With oMuestra
                .setTIPO_MUESTRA_ID = oTipo_analisis.getTIPO_MUESTRA_ID
                .setTIPO_ANALISIS_ID = oce_tipo_ensayo.getTIPO_ANALISIS_ID
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
                .setOBSERVACIONES_ENTREGA = txtob.Text
                .setCLIENTE_ID = cmbClientes.getPK_SALIDA
                .setCENTRO_ID = cmbCentro.BoundText
                .setREFERENCIA_CLIENTE = xA(i, ColsA.REFCLIENTE)
                ' Calculo de dias estimados
                Dim FechaEntrega As Date
                'M1356-I
                'FechaEntrega = DateAdd("d", oTipo_analisis.getDIAS_TRABAJO, fecha.value)
                FechaEntrega = xA(i, ColsA.fentrega)
                'M1356-F
                .setFECHA_PREV_FIN = Format(FechaEntrega, "yyyy-mm-dd")
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
                .setANALISIS_DUPLICADO = oce_tipo_ensayo.getDUPLICADO
'                .setPRODUCTO = txtDatos(2)
                If xA(i, ColsA.ENAC) <> "0" Then
                    .setPRODUCTO = xA(i, ColsA.producto)
                Else
                    .setPRODUCTO = ""
                End If

                If cmbPedido.Text <> "" Then
                    .setPEDIDO_ID = cmbPedido.BoundText
                End If
                .setREPLACEMENT_ID = 0
                .setENAC = oce_tipo_ensayo.getENAC
                .setNADCAP = oce_tipo_ensayo.getNADCAP
                muestra = .guardarMuestra
                .informar_precio_muestra muestra
            End With
            ' Informar observacion del tipo de ensayo si la tuviera
            indice = 1
            If oce_tipo_ensayo.getOBSERVACIONES <> "" Then
                With oDatos_especificos
                       .setMUESTRA_ID = muestra
                       .setBANO_ID = cmbbanos.getPK_SALIDA
                       .setTIPO_DATO_ID = 1
                       .setVALOR = oce_tipo_ensayo.getOBSERVACIONES
                       .setORDEN = indice
                       .Insertar
                       indice = indice + 1
                End With
            End If
            ' Datos específicos de la muestra
            Set rs = oTDA.Listado_por_tipo_analisis(oce_tipo_ensayo.getTIPO_ANALISIS_ID)
            If rs.RecordCount > 0 Then
                Do
                    With oDatos_especificos
                        .setMUESTRA_ID = muestra
                        .setBANO_ID = cmbbanos.getPK_SALIDA
                        .setTIPO_DATO_ID = rs(0)
                        If rs(0) = 28 Then ' Orden de compra
                            .setVALOR = txtDatos(0)
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
            With oce_recepcion
                .setNUMERO_RECEPCION = .getNUMERO_RECEPCION
                .setMUESTRA_ID = muestra
                .setTIPO_ENSAYO_ID = CLng(xA(i, ColsA.ID_TIPO_ENSAYO))
                .setDESIGNACION = xA(i, ColsA.DESIGNACION)
'                If chkSinEspecificar.value = Unchecked Then
'                    .setFECHA_PROCESADO_PIEZAS = Format(fprocesado.value, "yyyy-mm-dd")
'                Else
'                    .setFECHA_PROCESADO_PIEZAS = ""
'                End If
                'M1104-I
                If Trim(xA(i, ColsA.FPROCESO)) = "" Then
                    .setFECHA_PROCESADO_PIEZAS = "NULL"
                Else
                    .setFECHA_PROCESADO_PIEZAS = "'" & Format(xA(i, ColsA.FPROCESO), "yyyy-mm-dd") & "'"
                End If
                'M1104-F
                If oce_tipo_ensayo.getINCLUYE_ESPESOR = 1 Then
                    .setESPESOR = txtDatos(1)
                Else
                    .setESPESOR = "No requiere espesor."
                End If
                If oce_tipo_ensayo.getLOTE_PROBETAS = 1 Then
                    If cmbLote.getTEXTO <> "" Then
                        .setLOTE_PROBETA_ID = cmbLote.getPK_SALIDA
                    End If
                End If
                .setIDENTIFICACION_LABORATORIO = 0
                ' Insertar el equipo por defecto en la recepcion
'                If oce_tipo_ensayo.getEQUIPO_ID <> 0 Then
'                    .setMAQUINA = CStr(oce_tipo_ensayo.getEQUIPO_ID) & ";"
'                End If
'                .setCONDICIONES_AMBIENTALES = oce_tipo_ensayo.getCONDICIONES_AMBIENTALES
                .setCONDICIONES_AMBIENTALES = ""
                .setMATERIAL = ""
                .setDIMENSION = "0"
                .setMAQUINA = ""
'                If oce_tipo_ensayo.getINCLUYE_EQUIPO = 1 Then
                    Dim rsCe_Equipos As ADODB.Recordset
                    'M1137-I
                    'Set rsCe_Equipos = oce_tipo_ensayo.Equipos_Listado(oce_tipo_ensayo.getID_TIPO_ENSAYO)
                    Set rsCe_Equipos = oce_tipo_ensayo.Equipos_Listado_Recepcion(oce_tipo_ensayo.getID_TIPO_ENSAYO)
                    'M1137-F
                    If rsCe_Equipos.RecordCount > 0 Then
                        Do
                            .setMAQUINA = .getMAQUINA & rsCe_Equipos(0) & ";"
                            
                            oCE_EQUIPO.setMUESTRA_ID = muestra
                            oCE_EQUIPO.setORDEN = rsCe_Equipos(4)
                            oCE_EQUIPO.setEQUIPO_ID = rsCe_Equipos(0)
                            oCE_EQUIPO.setEN_INFORME = rsCe_Equipos(3)
                            oCE_EQUIPO.Insertar
                            
                            rsCe_Equipos.MoveNext
                        Loop Until rsCe_Equipos.EOF
                    End If
                    Set rsCe_Equipos = Nothing
'                End If
                'M0960-I Incluir reactivos en la recepcion
                Dim oReactivos As New clsCe_tipos_ensayos_botes_ex
                Dim rsReactivos As ADODB.Recordset
                Dim rExternos As String
                Dim rPropios As String
                Set rsReactivos = oReactivos.Listado(CLng(xA(i, ColsA.ID_TIPO_ENSAYO)))
                If rsReactivos.RecordCount > 0 Then
                    Do
                        If rsReactivos(1) = "E" Then
                            rExternos = rExternos & rsReactivos(0) & ";"
                        Else
                            rPropios = rPropios & rsReactivos(0) & ";"
                        End If
                        rsReactivos.MoveNext
                    Loop Until rsReactivos.EOF
                End If
                .setREACTIVOS = rExternos
                .setREACTIVOS_PROPIOS = rPropios
                'M0960-F
               .Insertar
            End With
            ' Resultados CE (campos PROBETA, DIMENSION) y calculamos la CANTIDAD de probetas
            Dim j As Integer
            Dim k As Integer
            Dim l As Integer
            Dim CANTIDAD_PROBETAS As Integer
            Dim CANTIDAD_POR_DESIGNACION As Integer
            CANTIDAD_PROBETAS = 0
            CANTIDAD_POR_DESIGNACION = 0
            With oCe_resultados
              For j = 0 To filasP
                If Not IsEmpty(xP(j, ColsP.DESIGNACION)) Then
                  If Trim(xP(j, ColsP.DESIGNACION)) <> "" Then
                    If xA(i, ColsA.DESIGNACION) = "TODAS" Or Trim(CStr(xP(j, ColsP.DESIGNACION))) = Trim(CStr(xA(i, ColsA.DESIGNACION))) Then
                       .setMUESTRA_ID = muestra
                       For k = 1 To CInt(xP(j, ColsP.RECIBIDAS))
                            .setPROBETA = k + CANTIDAD_PROBETAS
                            .setMATERIAL = Trim(CStr(xP(j, ColsP.MATERIAL)))
                            .setDIMENSION = Trim(CStr(xP(j, ColsP.DIMENSION)))
                            .setDESIGNACION = CStr(xP(j, ColsP.DESIGNACION))
                            If CInt(xP(j, ColsP.AREAS)) = 0 Then
                                .setAREA = 0
                                .Insertar
                            Else
                                For l = 1 To CInt(xP(j, ColsP.AREAS))
                                    .setAREA = l
                                    .Insertar
                                Next
                            End If
                       Next
                       CANTIDAD_POR_DESIGNACION = CANTIDAD_POR_DESIGNACION + CInt(xP(j, ColsP.RECIBIDAS))
                    End If
                    CANTIDAD_PROBETAS = CANTIDAD_PROBETAS + CInt(xP(j, ColsP.RECIBIDAS))
                  End If
                End If
              Next
            End With
            ' Informar la cantidad de probetas recibidas
            oce_recepcion.Informar_cantidad muestra, CANTIDAD_POR_DESIGNACION
            ' AIM
            Dim oAO As New clsAirbus_objetos
            Dim oMA As New clsMuestras_airbus
            With oAO
                If .Carga(TOBJETO.TOBJETO_BANO, cmbbanos.getPK_SALIDA) Then
                    oMA.setMUESTRA_ID = muestra
                    oMA.setENSAYO_ID = .getENSAYO_ID
                    oMA.setPROGRAMA_ID = .getPROGRAMA_ID
                    oMA.setFACILITY_ID = .getFACILITY_ID
                    oMA.setFLUID_ID = .getFLUID_ID
                    oMA.setSECTION_ID = .getSECTION_ID
                    oMA.Insertar True, True, True, True, True
                End If
            End With
          
          End If
         End If
        Next
        
        Me.MousePointer = 0
        MsgBox "La recepción se ha realizado correctamente. Proceda ahora a informar los datos de las probetas.", vbInformation, App.Title
        With frmCE_Recepcion_Detalle_Probetas
            .lDESIGNACION = ""
            .lProbetas = ""
            .lAreas = ""
            .lMaterial = ""
            .lDimensiones = ""
            For i = 0 To filasP
              If Not IsEmpty(xP(i, ColsP.DESIGNACION)) Then
               If Trim(xP(i, ColsP.DESIGNACION)) <> "" Then
                  .lDESIGNACION = .lDESIGNACION & Trim(xP(i, ColsP.DESIGNACION)) & ";"
                  .lProbetas = .lProbetas & CInt(xP(i, ColsP.RECIBIDAS)) & ";"
                  .lAreas = .lAreas & CInt(xP(i, ColsP.AREAS)) & ";"
                  .lMaterial = .lMaterial & Trim(xP(i, ColsP.MATERIAL)) & ";"
                  .lDimensiones = .lDimensiones & Trim(xP(i, ColsP.DIMENSION)) & ";"
               End If
              End If
            Next
            .NUMERO_RECEPCION = oce_recepcion.getNUMERO_RECEPCION
            .BANO = cmbbanos.getPK_SALIDA
            .CLIENTE_ID = cmbClientes.getPK_SALIDA
            .Show 1
        End With
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
    fecha = Date
'    fprocesado = Date
    fprocesado = "01-01-1900"
    cargar_producto

   On Error GoTo 0
   Exit Sub

Form_Load_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure Form_Load of Formulario frmCE_Recepcion"
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
    If txtDatos(0) = "" Then
        MsgBox "Informe la orden de compra.", vbExclamation, App.Title
        validar = False
        txtDatos(0).SetFocus
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
    If chkSinEspecificar.Value = Unchecked Then
        If Format(fprocesado.Value, "dd-mm-yyyy") = "01-01-1900" Then
            MsgBox "Debe indicar la fecha de procesado.", vbExclamation, App.Title
            validar = False
            fprocesado.SetFocus
            Exit Function
        Else
            If Year(fprocesado.Value) < Year(Date) - 1 Then
                MsgBox "Ojo, la fecha de procesado no parece correcta.", vbExclamation, App.Title
                validar = False
                fprocesado.SetFocus
                Exit Function
            End If
        End If
    End If
    ' Descripción del producto
'    If txtDatos(2).Visible = True Then
'        If txtDatos(2) = "" Then
'            MsgBox "Debe indicar la desripción del producto.", vbExclamation, App.Title
'            validar = False
'            txtDatos(2).SetFocus
'            Exit Function
'        End If
'    End If
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
            If xP(i, ColsP.DESIGNACION) = "" And Trim(xP(i, ColsP.RECIBIDAS)) <> "" Then
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
        If Not IsEmpty(xA(i, ColsA.tipo_ensayo)) Then
         If Trim(xA(i, ColsA.tipo_ensayo)) <> "" Then
            algo = True
            If IsEmpty(xA(i, ColsA.REFCLIENTE)) Then
                ref_cliente = True
            End If
            If IsEmpty(xA(i, ColsA.DESIGNACION)) Then
                DESIG = True
            End If
         End If
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
    Dim oce_tipo_ensayo As New clsCe_tipos_ensayos
    ' Validar el Espesor y el lote de probetas
    For i = 0 To filasA
        If Not IsEmpty(xA(i, ColsA.tipo_ensayo)) Then
         If Trim(CStr(xA(i, ColsA.tipo_ensayo))) <> "" Then
            If oce_tipo_ensayo.Carga(xA(i, ColsA.ID_TIPO_ENSAYO)) Then
               If oce_tipo_ensayo.getINCLUYE_ESPESOR = 1 Then
                   If txtDatos(1) = "" Then
                       MsgBox "Debe indicar el espesor.", vbExclamation, App.Title
                       validar = False
                       txtDatos(1).SetFocus
                       Exit Function
                   End If
               End If
               If oce_tipo_ensayo.getLOTE_PROBETAS = 1 Then
                   If cmbLote.getPK_SALIDA = 0 Then
                       MsgBox "Debe indicar el lote de probetas.", vbExclamation, App.Title
                       validar = False
                       cmbLote.SetFocus
                       Exit Function
                   End If
               End If
            End If
         End If
        End If
    Next
    ' Validar que los ensayos ENAC tengan informado la descripcion del producto
    For i = 0 To filasA
        If Not IsEmpty(xA(i, ColsA.tipo_ensayo)) Then
         If Trim(CStr(xA(i, ColsA.tipo_ensayo))) <> "" Then
          If xA(i, ColsA.ENAC) <> "0" Then
           If xA(i, ColsA.producto) = "" Then
             MsgBox "Debe indicar la Descripción del Producto para los ensayos ENAC.", vbExclamation, App.Title
             validar = False
             Exit Function
           End If
          End If
         End If
        End If
    Next

   On Error GoTo 0
   Exit Function

validar_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure validar of Formulario frmCE_Recepcion"
End Function
Private Sub cargar_combos()
    cargar_clientes
    cargar_combo cmbCentro, New clsCentros
    llenar_combo cmbbanos, New clsBanos, 0, frmBANO_Detalle, " ANULADO = 0 "
    cargar_combo cmbenvases, New clsformatos
    cargar_combo cmbentregada, New clsEntidades_Entrega
    cargar_combo cmbrealizada, New clsEntidades_muestreo
    llenar_combo cmbLote, New clsCe_lotes_probetas, 0, frmCE_Lote_Probeta, ""
    cmbLote.desactivar
End Sub
Private Sub cargar_ficha(BANO As Long, FICHA As Long)
   On Error GoTo cargar_ficha_Error

    inicializar_grid
    Dim oCe_Ficha As New clsCe_ficha
    If oCe_Ficha.Carga(FICHA) Then
        txtproceso = oCe_Ficha.getPROCESO
    End If
    ' Recuperamos los datos de la ficha de proceso
    Dim oCe_bano_probetas As New clsCe_banos_probetas
    Dim i As Integer
    i = 0
    Dim rs As ADODB.Recordset
    Set rs = oCe_bano_probetas.Listado(BANO)
    If rs.RecordCount > 0 Then
        Do
            xP(i, ColsP.DESIGNACION) = CStr(rs(0))
            xP(i, ColsP.MATERIAL) = CStr(rs(1))
            xP(i, ColsP.DES_PRODUCTO) = CStr(rs(2))
            xP(i, ColsP.DIMENSION) = CStr(rs(3))
            xP(i, ColsP.nProbetas) = CStr(rs(4))
            xP(i, ColsP.AREAS) = CStr(rs(5))
            i = i + 1
            rs.MoveNext
        Loop Until rs.EOF
    End If
    Dim oCe_bano_ensayos As New clsCe_banos_ensayos
    i = 0
    txtDatos(1) = "µm"
    txtDatos(1).Enabled = False
    cmbLote.limpiar
    cmbLote.desactivar
    Set rs = oCe_bano_ensayos.Listado(BANO)
    If rs.RecordCount > 0 Then
        Do
            xA(i, ColsA.tipo_ensayo) = CStr(rs(0))
            xA(i, ColsA.NORMA) = CStr(rs(1))
            xA(i, ColsA.DESIGNACION) = CStr(rs(2))
            xA(i, ColsA.ID_TIPO_ENSAYO) = CStr(rs(3))
            xA(i, ColsA.ENAC) = CStr(rs(6))
            'M1356-I
            xA(i, ColsA.fentrega) = calcularFechaFinalizacion(fecha.Value, CInt(rs(7)))
            'M1356-F
            If rs(4) = 1 Then
                txtDatos(1).Enabled = True
            End If
            If rs(5) = 1 Then
                cmbLote.activar
            End If
            ' Buscar la Descripción del producto
            Dim j As Integer
            For j = 0 To filasP
                If Not IsEmpty(xP(j, ColsP.DESIGNACION)) Then
                    If UCase(Trim(CStr(rs(2)))) = UCase(Trim(xP(j, ColsP.DESIGNACION))) Then
                        xA(i, ColsA.producto) = xP(j, ColsP.DES_PRODUCTO)
                    End If
                End If
            Next
            i = i + 1
            rs.MoveNext
        Loop Until rs.EOF
    End If
    cargar_combo_tipos_ensayos FICHA
'    cargar_combo_probetas
    cargar_producto
    gridP.Refresh
    gridA.Refresh

   On Error GoTo 0
   Exit Sub

cargar_ficha_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure cargar_ficha of Formulario frmCE_Recepcion"
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
Public Sub cargar_banos()
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

Private Sub fprocesado_Change()
    informar_fechas_procesado
    calcular_referencias
End Sub

Private Sub gridA_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button And vbRightButton Then
        If Not IsEmpty(xA(gridA.Row, ColsA.ID_TIPO_ENSAYO)) Then
            frmCE_Tipo_Ensayo.PK = CLng(xA(gridA.Row, ColsA.ID_TIPO_ENSAYO))
            frmCE_Tipo_Ensayo.Show 1
        End If
    End If
End Sub

Private Sub gridP_KeyPress(KeyAscii As Integer)
    calcular_referencias
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

Private Sub txtdatos_GotFocus(Index As Integer)
    txtDatos(Index).BackColor = &H80FFFF
    txtDatos(Index).SelStart = 0
    txtDatos(Index).SelLength = Len(txtDatos(Index))
End Sub
Private Sub txtdatos_LostFocus(Index As Integer)
    txtDatos(Index).BackColor = vbWhite
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

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure inicializar_grid of Formulario frmCE_Recepcion_Nuevo2"
End Sub

Private Sub cargar_combo_tipos_ensayos(FICHA As Long)
    Dim rs As ADODB.Recordset
    Dim ote As New clsCe_tipos_ensayos
'    Set rs = ote.Listado(ficha)
        'M1147-I
        'Set rs = ote.Listado("", "", True)
    Set rs = ote.Listado("", "", True, False)
        'M1147-F
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
'    gridA.Refresh
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

Private Sub cargar_pedidos(cliente As Long, fecha As Date)
    Dim oPedido As New clsClientes_pedidos
    Set cmbPedido.RowSource = oPedido.Listado_en_fecha(CInt(cliente), CStr(fecha))
    cmbPedido.ListField = "CODIGO_LARGO"
    cmbPedido.DataField = "ID_PEDIDO"
    cmbPedido.BoundColumn = "ID_PEDIDO"
End Sub
Private Sub calcular_referencias()
    Dim i As Integer
    Dim j As Integer
    Dim CANTIDAD As Integer
    Dim ref As String
    Dim fecha As String
    ' Calculamos la fecha de procesado de las piezas
    If Format(fprocesado.Value, "dd-mm-yyyy") <> "01-01-1900" Then
        If chkFechaCompleta.Value = Checked Then
            fecha = Format(fprocesado, "dd/MMMM/yy")
        Else
            fecha = Format(fprocesado, "MMMM/yy")
        End If
    End If
    ' Asignamos las ref. correspondientes
    Dim cantidad_distinta As Boolean
    cantidad_distinta = False
    For i = 0 To filasA
      If Not IsEmpty(xA(i, ColsA.ID_TIPO_ENSAYO)) Then
        If Trim(xA(i, ColsA.ID_TIPO_ENSAYO)) <> "" Then
            ' Contamos la cantidad de probetas recibidas
            CANTIDAD = 0
            For j = 0 To filasP
                If Not IsEmpty(xP(j, ColsP.DESIGNACION)) Then
                    If UCase(Trim(xA(i, ColsA.DESIGNACION))) = UCase(Trim(xP(j, ColsP.DESIGNACION))) Or _
                       UCase(Trim(xA(i, ColsA.DESIGNACION))) = "TODAS" Then
                        If IsNumeric(xP(j, ColsP.RECIBIDAS)) Then
                            CANTIDAD = CANTIDAD + CInt(xP(j, ColsP.RECIBIDAS))
                            ' Comparar las introducidas con las recibidas
                            ' Si son distintas, incluirlo en las observaciones
                            If IsNumeric(xP(j, ColsP.nProbetas)) Then
                                If CInt(xP(j, ColsP.nProbetas)) <> CInt(xP(j, ColsP.RECIBIDAS)) Then
                                    cantidad_distinta = True
                                End If
                            End If
                        Else
                            cantidad_distinta = False
                        End If
                    End If
                End If
            Next
            ' La referencia del cliente se compone de :
            ' NºPROBETAS + DESIGNACION + SOLUCION + BAÑO + MES PROCESO (O FECHA COMPLETA)
'            ref = CStr(cantidad & " PROBETAS: " & UCase(Trim(xA(i, ColsA.DESIGNACION))) & " " & _
'                  txtsolucion & " BAÑO: " & cmbbanos.getTEXTO)
            ' Se deja solo la cantidad a petición de Mari Cruz (17/01/2017)
            ref = CStr(CANTIDAD & " PROBETAS: " & UCase(Trim(xA(i, ColsA.DESIGNACION))))
            xA(i, ColsA.REFCLIENTE) = ref & " " & UCase(fecha)
        End If
      End If
    Next
    If Not cantidad_distinta Then
        txtob = ""
    Else
        txtob = "El número de probetas recibidas es distinto al de la norma."
    End If
    gridA.Refresh
End Sub
Private Sub informar_fechas_procesado()
    Dim i As Integer
    For i = 0 To filasA
      If Not IsEmpty(xA(i, ColsA.DESIGNACION)) Then
        If chkSinEspecificar.Value = Checked Then
            xA(i, ColsA.FPROCESO) = ""
        Else
            xA(i, ColsA.FPROCESO) = Format(fprocesado.Value, "dd-mm-yyyy")
        End If
      End If
    Next
    gridP.Refresh
End Sub
Private Sub cargar_producto()
    xProducto.Clear
    xProducto.ReDim 1, 1, 1, 1
    xProducto(1, 1) = " "
    Set tProducto.Array = xProducto
    tProducto.Refresh
    Dim i As Integer
    Dim join As String
    join = ""
    For i = 0 To filasA
        If Not IsEmpty(xA(i, ColsA.ID_TIPO_ENSAYO)) Then
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
    consulta = "SELECT DISTINCT DESCRIPCION " & _
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

