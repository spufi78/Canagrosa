VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{CEA3B7F6-2847-4E5E-A551-DB7A62489D44}#46.0#0"; "miCombo.ocx"
Object = "{77DDF307-D82B-4757-8B3A-106EC9D75D4B}#8.0#0"; "tdbg8.ocx"
Begin VB.Form frmCE_Recepcion_Probetas 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Recepción de Control de Eficacia"
   ClientHeight    =   9765
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   15555
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "frmCE_Recepcion_Probetas_Nuevo.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9765
   ScaleWidth      =   15555
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame frmCOC 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Datos COC"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3525
      Left            =   4770
      TabIndex        =   18
      Top             =   3060
      Visible         =   0   'False
      Width           =   6405
      Begin VB.TextBox txtCOC 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   330
         Index           =   1
         Left            =   1305
         TabIndex        =   24
         Top             =   675
         Width           =   1470
      End
      Begin VB.CommandButton cmdInformarCOC 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Informar"
         Height          =   510
         Left            =   4005
         Style           =   1  'Graphical
         TabIndex        =   23
         Top             =   2925
         Width           =   1140
      End
      Begin VB.CommandButton cmdCerrarCOC 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Salir"
         Height          =   510
         Left            =   5175
         Style           =   1  'Graphical
         TabIndex        =   22
         Top             =   2925
         Width           =   1050
      End
      Begin VB.TextBox txtCOC 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   330
         Index           =   0
         Left            =   1305
         TabIndex        =   20
         Top             =   1755
         Width           =   1470
      End
      Begin VB.TextBox txtCOC 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   330
         Index           =   2
         Left            =   1305
         TabIndex        =   19
         Top             =   2115
         Width           =   1470
      End
      Begin pryCombo.miCombo cmbDimension 
         Height          =   375
         Index           =   1
         Left            =   1305
         TabIndex        =   21
         Top             =   1395
         Width           =   4785
         _ExtentX        =   8440
         _ExtentY        =   661
      End
      Begin pryCombo.miCombo cmbDimension 
         Height          =   375
         Index           =   0
         Left            =   1305
         TabIndex        =   25
         Top             =   1035
         Width           =   4785
         _ExtentX        =   8440
         _ExtentY        =   661
      End
      Begin MSComCtl2.DTPicker fecha 
         Height          =   330
         Left            =   1305
         TabIndex        =   26
         Top             =   270
         Width           =   1470
         _ExtentX        =   2593
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
         Format          =   61276161
         CurrentDate     =   38000
         MinDate         =   2
      End
      Begin VB.Label lblSet 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C0C0&
         Caption         =   "SET : Insertar las dos dimensiones"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   285
         Left            =   135
         TabIndex        =   35
         Top             =   2565
         Visible         =   0   'False
         Width           =   6090
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Ej. 0201"
         ForeColor       =   &H00FF0000&
         Height          =   195
         Index           =   4
         Left            =   2880
         TabIndex        =   34
         Top             =   2205
         Width           =   585
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Ej. 3124E"
         ForeColor       =   &H00FF0000&
         Height          =   195
         Index           =   3
         Left            =   2880
         TabIndex        =   33
         Top             =   765
         Width           =   690
      End
      Begin VB.Label Label1 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Dimension 1"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   4
         Left            =   135
         TabIndex        =   32
         Top             =   1080
         Width           =   1215
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Código"
         Height          =   195
         Index           =   9
         Left            =   135
         TabIndex        =   31
         Top             =   720
         Width           =   495
      End
      Begin VB.Label Label1 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Dimension 2"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   1
         Left            =   135
         TabIndex        =   30
         Top             =   1440
         Width           =   1215
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Fecha"
         Height          =   195
         Index           =   0
         Left            =   135
         TabIndex        =   29
         Top             =   315
         Width           =   450
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Nº Proceso"
         Height          =   195
         Index           =   1
         Left            =   135
         TabIndex        =   28
         Top             =   1800
         Width           =   810
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Cadena"
         Height          =   195
         Index           =   2
         Left            =   135
         TabIndex        =   27
         Top             =   2160
         Width           =   555
      End
   End
   Begin VB.CommandButton cmdCOC 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Informar COC"
      Height          =   870
      Left            =   1260
      Picture         =   "frmCE_Recepcion_Probetas_Nuevo.frx":2AFA
      Style           =   1  'Graphical
      TabIndex        =   17
      Top             =   8865
      Width           =   1140
   End
   Begin VB.CheckBox chkIL 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Identificación aportada por Canagrosa"
      Height          =   240
      Left            =   3870
      TabIndex        =   16
      Top             =   9495
      Width           =   5685
   End
   Begin VB.CommandButton cmdborrar 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Borrar Probeta"
      Height          =   870
      Left            =   90
      Picture         =   "frmCE_Recepcion_Probetas_Nuevo.frx":33C4
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   8865
      Width           =   1140
   End
   Begin VB.CommandButton cmdok 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Aceptar"
      Height          =   870
      Left            =   13320
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   8865
      Width           =   1050
   End
   Begin VB.TextBox txtcantidad 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   12
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   375
      Left            =   10035
      TabIndex        =   12
      Top             =   270
      Width           =   1005
   End
   Begin VB.CommandButton cmdCopiar 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      Caption         =   "Copiar"
      Height          =   285
      Index           =   1
      Left            =   13050
      Style           =   1  'Graphical
      TabIndex        =   10
      ToolTipText     =   "Copia la dimensión seleccionada en todas las probetas siguientes"
      Top             =   1035
      Width           =   690
   End
   Begin VB.CommandButton cmdCopiar 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      Caption         =   "Copiar"
      Height          =   285
      Index           =   0
      Left            =   10620
      Style           =   1  'Graphical
      TabIndex        =   9
      ToolTipText     =   "Copia el material seleccionado en todas las probetas siguientes"
      Top             =   1035
      Width           =   690
   End
   Begin VB.CommandButton CMDINFORMAR 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Informar"
      Height          =   330
      Index           =   1
      Left            =   4995
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   8820
      Width           =   690
   End
   Begin VB.TextBox txtDatos 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Height          =   330
      Index           =   3
      Left            =   3015
      TabIndex        =   7
      Top             =   8820
      Width           =   1965
   End
   Begin VB.TextBox txtDatos 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Height          =   330
      Index           =   2
      Left            =   5715
      TabIndex        =   6
      Top             =   8820
      Width           =   2325
   End
   Begin VB.CommandButton CMDINFORMAR 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Informar"
      Height          =   330
      Index           =   0
      Left            =   8100
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   8820
      Width           =   735
   End
   Begin VB.CommandButton cmdSalir 
      BackColor       =   &H00E0E0E0&
      Caption         =   "ESC-Salir"
      Height          =   870
      Left            =   14415
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   8865
      Width           =   1050
   End
   Begin TrueDBGrid80.TDBGrid grid 
      Height          =   8040
      Left            =   45
      TabIndex        =   0
      Top             =   765
      Width           =   15450
      _ExtentX        =   27252
      _ExtentY        =   14182
      _LayoutType     =   4
      _RowHeight      =   -2147483647
      _WasPersistedAsPixels=   0
      Columns(0)._VlistStyle=   0
      Columns(0)._MaxComboItems=   5
      Columns(0).Caption=   "Designación"
      Columns(0).DataField=   ""
      Columns(0).NumberFormat=   "General Number"
      Columns(0).ExternalEditor=   "TDBDate1"
      Columns(0).ExternalEditor.vt=   8
      Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(1)._VlistStyle=   0
      Columns(1)._MaxComboItems=   5
      Columns(1).Caption=   "Probeta"
      Columns(1).DataField=   ""
      Columns(1).NumberFormat=   "General Number"
      Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(2)._VlistStyle=   0
      Columns(2)._MaxComboItems=   5
      Columns(2).Caption=   "Area"
      Columns(2).DataField=   ""
      Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(3)._VlistStyle=   0
      Columns(3)._MaxComboItems=   5
      Columns(3).Caption=   "Ident. Canagrosa"
      Columns(3).DataField=   ""
      Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(4)._VlistStyle=   0
      Columns(4)._MaxComboItems=   5
      Columns(4).Caption=   "Ident. Cliente"
      Columns(4).DataField=   ""
      Columns(4)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(5)._VlistStyle=   0
      Columns(5)._MaxComboItems=   5
      Columns(5).Caption=   "Material"
      Columns(5).DataField=   ""
      Columns(5).ExternalEditor=   "TDBDate1"
      Columns(5).ExternalEditor.vt=   8
      Columns(5)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(6)._VlistStyle=   0
      Columns(6)._MaxComboItems=   5
      Columns(6).Caption=   "Dimensión mm"
      Columns(6).DataField=   ""
      Columns(6)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(7)._VlistStyle=   0
      Columns(7)._MaxComboItems=   5
      Columns(7).Caption=   "C.A."
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
      Splits(0)._ColumnProps(1)=   "Column(0).Width=2461"
      Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
      Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=2381"
      Splits(0)._ColumnProps(4)=   "Column(0)._EditAlways=0"
      Splits(0)._ColumnProps(5)=   "Column(0)._ColStyle=1"
      Splits(0)._ColumnProps(6)=   "Column(0).Order=1"
      Splits(0)._ColumnProps(7)=   "Column(1).Width=1455"
      Splits(0)._ColumnProps(8)=   "Column(1).DividerColor=0"
      Splits(0)._ColumnProps(9)=   "Column(1)._WidthInPix=1376"
      Splits(0)._ColumnProps(10)=   "Column(1)._EditAlways=0"
      Splits(0)._ColumnProps(11)=   "Column(1)._ColStyle=1"
      Splits(0)._ColumnProps(12)=   "Column(1).Order=2"
      Splits(0)._ColumnProps(13)=   "Column(2).Width=1323"
      Splits(0)._ColumnProps(14)=   "Column(2).DividerColor=0"
      Splits(0)._ColumnProps(15)=   "Column(2)._WidthInPix=1244"
      Splits(0)._ColumnProps(16)=   "Column(2)._EditAlways=0"
      Splits(0)._ColumnProps(17)=   "Column(2)._ColStyle=1"
      Splits(0)._ColumnProps(18)=   "Column(2).Order=3"
      Splits(0)._ColumnProps(19)=   "Column(3).Width=4815"
      Splits(0)._ColumnProps(20)=   "Column(3).DividerColor=0"
      Splits(0)._ColumnProps(21)=   "Column(3)._WidthInPix=4736"
      Splits(0)._ColumnProps(22)=   "Column(3)._EditAlways=0"
      Splits(0)._ColumnProps(23)=   "Column(3)._ColStyle=1"
      Splits(0)._ColumnProps(24)=   "Column(3).Order=4"
      Splits(0)._ColumnProps(25)=   "Column(3).AutoDropDown=1"
      Splits(0)._ColumnProps(26)=   "Column(3).DropDownList=1"
      Splits(0)._ColumnProps(27)=   "Column(3).AutoCompletion=1"
      Splits(0)._ColumnProps(28)=   "Column(4).Width=5424"
      Splits(0)._ColumnProps(29)=   "Column(4).DividerColor=0"
      Splits(0)._ColumnProps(30)=   "Column(4)._WidthInPix=5345"
      Splits(0)._ColumnProps(31)=   "Column(4)._EditAlways=0"
      Splits(0)._ColumnProps(32)=   "Column(4)._ColStyle=1"
      Splits(0)._ColumnProps(33)=   "Column(4).Order=5"
      Splits(0)._ColumnProps(34)=   "Column(4).AutoDropDown=1"
      Splits(0)._ColumnProps(35)=   "Column(4).AutoCompletion=1"
      Splits(0)._ColumnProps(36)=   "Column(5).Width=4392"
      Splits(0)._ColumnProps(37)=   "Column(5).DividerColor=0"
      Splits(0)._ColumnProps(38)=   "Column(5)._WidthInPix=4313"
      Splits(0)._ColumnProps(39)=   "Column(5)._EditAlways=0"
      Splits(0)._ColumnProps(40)=   "Column(5)._ColStyle=1"
      Splits(0)._ColumnProps(41)=   "Column(5).Order=6"
      Splits(0)._ColumnProps(42)=   "Column(6).Width=4286"
      Splits(0)._ColumnProps(43)=   "Column(6).DividerColor=0"
      Splits(0)._ColumnProps(44)=   "Column(6)._WidthInPix=4207"
      Splits(0)._ColumnProps(45)=   "Column(6)._EditAlways=0"
      Splits(0)._ColumnProps(46)=   "Column(6)._ColStyle=1"
      Splits(0)._ColumnProps(47)=   "Column(6).Order=7"
      Splits(0)._ColumnProps(48)=   "Column(7).Width=265"
      Splits(0)._ColumnProps(49)=   "Column(7).DividerColor=0"
      Splits(0)._ColumnProps(50)=   "Column(7)._WidthInPix=185"
      Splits(0)._ColumnProps(51)=   "Column(7)._EditAlways=0"
      Splits(0)._ColumnProps(52)=   "Column(7)._ColStyle=1"
      Splits(0)._ColumnProps(53)=   "Column(7).Order=8"
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
      Caption         =   "Identificación de las probetas"
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
      _StyleDefs(36)  =   "Splits(0).Columns(0).Style:id=24,.parent=11,.alignment=2,.bgcolor=&HC1FFFF&"
      _StyleDefs(37)  =   ":id=24,.fgcolor=&HFF&,.locked=0,.bold=-1,.fontsize=975,.italic=0,.underline=0"
      _StyleDefs(38)  =   ":id=24,.strikethrough=0,.charset=0"
      _StyleDefs(39)  =   ":id=24,.fontname=MS Sans Serif"
      _StyleDefs(40)  =   "Splits(0).Columns(0).HeadingStyle:id=21,.parent=12"
      _StyleDefs(41)  =   "Splits(0).Columns(0).FooterStyle:id=22,.parent=13"
      _StyleDefs(42)  =   "Splits(0).Columns(0).EditorStyle:id=23,.parent=15,.bold=0,.fontsize=975"
      _StyleDefs(43)  =   ":id=23,.italic=0,.underline=0,.strikethrough=0,.charset=0"
      _StyleDefs(44)  =   ":id=23,.fontname=MS Sans Serif"
      _StyleDefs(45)  =   "Splits(0).Columns(1).Style:id=32,.parent=11,.alignment=2,.bgcolor=&HC1FFFF&"
      _StyleDefs(46)  =   ":id=32,.fgcolor=&HFF&,.locked=0,.bold=-1,.fontsize=975,.italic=0,.underline=0"
      _StyleDefs(47)  =   ":id=32,.strikethrough=0,.charset=0"
      _StyleDefs(48)  =   ":id=32,.fontname=MS Sans Serif"
      _StyleDefs(49)  =   "Splits(0).Columns(1).HeadingStyle:id=29,.parent=12"
      _StyleDefs(50)  =   "Splits(0).Columns(1).FooterStyle:id=30,.parent=13"
      _StyleDefs(51)  =   "Splits(0).Columns(1).EditorStyle:id=31,.parent=15"
      _StyleDefs(52)  =   "Splits(0).Columns(2).Style:id=62,.parent=11,.alignment=2,.bgcolor=&HC1FFFF&"
      _StyleDefs(53)  =   ":id=62,.fgcolor=&HFF&,.bold=-1,.fontsize=975,.italic=0,.underline=0"
      _StyleDefs(54)  =   ":id=62,.strikethrough=0,.charset=0"
      _StyleDefs(55)  =   ":id=62,.fontname=MS Sans Serif"
      _StyleDefs(56)  =   "Splits(0).Columns(2).HeadingStyle:id=59,.parent=12"
      _StyleDefs(57)  =   "Splits(0).Columns(2).FooterStyle:id=60,.parent=13"
      _StyleDefs(58)  =   "Splits(0).Columns(2).EditorStyle:id=61,.parent=15"
      _StyleDefs(59)  =   "Splits(0).Columns(3).Style:id=36,.parent=11,.alignment=2,.bgcolor=&HC1FFFF&"
      _StyleDefs(60)  =   ":id=36,.fgcolor=&HFF&,.bold=-1,.fontsize=975,.italic=0,.underline=0"
      _StyleDefs(61)  =   ":id=36,.strikethrough=0,.charset=0"
      _StyleDefs(62)  =   ":id=36,.fontname=MS Sans Serif"
      _StyleDefs(63)  =   "Splits(0).Columns(3).HeadingStyle:id=33,.parent=12"
      _StyleDefs(64)  =   "Splits(0).Columns(3).FooterStyle:id=34,.parent=13"
      _StyleDefs(65)  =   "Splits(0).Columns(3).EditorStyle:id=35,.parent=15"
      _StyleDefs(66)  =   "Splits(0).Columns(4).Style:id=54,.parent=11,.alignment=2"
      _StyleDefs(67)  =   "Splits(0).Columns(4).HeadingStyle:id=51,.parent=12"
      _StyleDefs(68)  =   "Splits(0).Columns(4).FooterStyle:id=52,.parent=13"
      _StyleDefs(69)  =   "Splits(0).Columns(4).EditorStyle:id=53,.parent=15"
      _StyleDefs(70)  =   "Splits(0).Columns(5).Style:id=28,.parent=11,.alignment=2"
      _StyleDefs(71)  =   "Splits(0).Columns(5).HeadingStyle:id=25,.parent=12"
      _StyleDefs(72)  =   "Splits(0).Columns(5).FooterStyle:id=26,.parent=13"
      _StyleDefs(73)  =   "Splits(0).Columns(5).EditorStyle:id=27,.parent=15"
      _StyleDefs(74)  =   "Splits(0).Columns(6).Style:id=58,.parent=11,.alignment=2"
      _StyleDefs(75)  =   "Splits(0).Columns(6).HeadingStyle:id=55,.parent=12"
      _StyleDefs(76)  =   "Splits(0).Columns(6).FooterStyle:id=56,.parent=13"
      _StyleDefs(77)  =   "Splits(0).Columns(6).EditorStyle:id=57,.parent=15"
      _StyleDefs(78)  =   "Splits(0).Columns(7).Style:id=66,.parent=11,.alignment=2"
      _StyleDefs(79)  =   "Splits(0).Columns(7).HeadingStyle:id=63,.parent=12"
      _StyleDefs(80)  =   "Splits(0).Columns(7).FooterStyle:id=64,.parent=13"
      _StyleDefs(81)  =   "Splits(0).Columns(7).EditorStyle:id=65,.parent=15"
      _StyleDefs(82)  =   "Named:id=37:Normal"
      _StyleDefs(83)  =   ":id=37,.parent=0,.alignment=3,.bgcolor=&H80000018&,.bold=0,.fontsize=825"
      _StyleDefs(84)  =   ":id=37,.italic=0,.underline=0,.strikethrough=0,.charset=0"
      _StyleDefs(85)  =   ":id=37,.fontname=MS Sans Serif"
      _StyleDefs(86)  =   "Named:id=38:Heading"
      _StyleDefs(87)  =   ":id=38,.parent=37,.valignment=2,.bgcolor=&H80000004&,.fgcolor=&H80000012&"
      _StyleDefs(88)  =   ":id=38,.wraptext=-1,.appearance=0,.ellipsis=0"
      _StyleDefs(89)  =   "Named:id=39:Footing"
      _StyleDefs(90)  =   ":id=39,.parent=37,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(91)  =   "Named:id=40:Selected"
      _StyleDefs(92)  =   ":id=40,.parent=37,.bgcolor=&H8080FF&,.fgcolor=&H0&,.bold=0,.fontsize=975"
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
   Begin VB.Label lblCerrada 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   345
      Left            =   7200
      TabIndex        =   15
      Top             =   135
      Width           =   1950
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0C0&
      Caption         =   "Para almacenar una probeta, debe rellenar los campos señalados en amarillo."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   285
      Left            =   2745
      TabIndex        =   14
      Top             =   9225
      Width           =   8790
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Número de Probetas Recibidas"
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
      Height          =   690
      Index           =   0
      Left            =   9225
      TabIndex        =   11
      Top             =   0
      Width           =   2805
   End
   Begin VB.Label lbltitulo 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "Datos de las Probetas del Control de Eficacia"
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
      TabIndex        =   4
      Top             =   90
      Width           =   4755
   End
   Begin VB.Image imagen 
      Height          =   480
      Left            =   14895
      Picture         =   "frmCE_Recepcion_Probetas_Nuevo.frx":3C8E
      Top             =   90
      Width           =   480
   End
   Begin VB.Label lblsubtitulo 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "Indique los datos de las probetas de los distintos ensayos de Eficacia"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   90
      TabIndex        =   3
      Top             =   360
      Width           =   4875
   End
   Begin VB.Shape fondo 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   735
      Left            =   0
      Top             =   0
      Width           =   15525
   End
End
Attribute VB_Name = "frmCE_Recepcion_Probetas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public PK_MUESTRA As Long
Dim x As New XArrayDB
Const Col As Integer = 7
Const filas As Integer = 1000
Private Enum COLS
    DESIGNACION = 0
    PROBETA = 1
    AREA = 2
    IDEN_CANAGROSA = 3
    IDEN_CLIENTE = 4
    material = 5
    dimension = 6
    CA = 7
End Enum

Private Sub cmdBorrar_Click()
    On Error Resume Next
    Dim i As Integer
    For i = 0 To Col
        grid.SelBookmarks.Add grid.Bookmark
        x(grid.Bookmark, i) = ""
        grid.SelBookmarks.Remove 0
    Next
    If grid.Row > 0 Then
        grid.Row = grid.Row - 1
    End If
    grid.Refresh
    grid.SetFocus
    contar_probetas
End Sub

Private Sub cmdCerrarCOC_Click()
    frmCOC.visible = False
End Sub

Private Sub cmdCOC_Click()
    frmCOC.visible = True
End Sub

Private Sub cmdCopiar_Click(Index As Integer)
    Dim i As Integer
   On Error GoTo cmdCopiar_Click_Error
    Dim material As String
    Dim dimension As String
    material = x(grid.Bookmark, COLS.material)
    dimension = x(grid.Bookmark, COLS.dimension)
    For i = grid.Bookmark To filas
     If Not IsEmpty(x(i, COLS.DESIGNACION)) Then
        Select Case Index
        Case 0
            x(i, COLS.material) = material
        Case 1
            x(i, COLS.dimension) = dimension
        End Select
     End If
    Next
    grid.Refresh

   On Error GoTo 0
   Exit Sub

cmdCopiar_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure cmdCopiar_Click of Formulario frmCE_Recepcion_Probetas"
End Sub

Private Sub cmdInformar_Click(Index As Integer)
    Select Case Index
    Case 0
        If txtDatos(2) = "" Then
            MsgBox "Indique el sufijo para generar.", vbExclamation, App.Title
            txtDatos(2).SetFocus
            Exit Sub
        End If
    Case 1
        If txtDatos(3) = "" Then
            MsgBox "Indique el sufijo para generar.", vbExclamation, App.Title
            txtDatos(3).SetFocus
            Exit Sub
        End If
    End Select
    Dim i As Integer
    For i = 1 To filas
     If Not IsEmpty(x(i, COLS.DESIGNACION)) Then
        Select Case Index
        Case 0
            x(i, COLS.IDEN_CLIENTE) = txtDatos(2) & "-" & i
        Case 1
            x(i, COLS.IDEN_CANAGROSA) = txtDatos(3) & "-" & i
        End Select
      End If
    Next
    grid.Refresh
End Sub

Private Sub cmdok_Click()
   On Error GoTo cmdok_Click_Error

'    If validar Then
        Me.MousePointer = 11
        Dim oCe_resultados As New clsCe_resultados
        Dim probetas As String
        Dim MUESTRA As Long
        For i = 1 To filas
           If Not IsEmpty(x(i, COLS.DESIGNACION)) And _
              Not IsEmpty(x(i, COLS.PROBETA)) And _
              Not IsEmpty(x(i, COLS.AREA)) And _
              Not IsEmpty(x(i, COLS.IDEN_CANAGROSA)) Then
             If CStr(x(i, COLS.IDEN_CANAGROSA)) <> "" Then
                With oCe_resultados
                  .setIDENTIFICACION_CLIENTE = x(i, COLS.IDEN_CLIENTE)
                  .setIDENTIFICACION_CANAGROSA = x(i, COLS.IDEN_CANAGROSA)
                  .setMATERIAL = x(i, COLS.material)
                  .setDIMENSION = x(i, COLS.dimension)
                  .setCRITERIO_ACEPTACION = x(i, COLS.CA)
                  .Modificar_datos_muestra PK_MUESTRA, CStr(x(i, COLS.DESIGNACION)), CInt(x(i, COLS.PROBETA)), CInt(x(i, COLS.AREA))
                  probetas = probetas & "," & x(i, COLS.PROBETA)
                End With
             End If
           End If
        Next
        oCe_resultados.Eliminar_no_utilizadas PK_MUESTRA, Right(probetas, Len(probetas) - 1)
        Dim oce_recepcion As New clsCe_recepcion
        oce_recepcion.Informar_cantidad PK_MUESTRA, CInt(txtcantidad)
        oce_recepcion.Informar_Identificacion_Laboratorio PK_MUESTRA, chkIL.Value
        
        MsgBox "Datos almacenados correctamente", vbInformation + vbOKOnly, App.Title
        Unload Me
'    End If

   On Error GoTo 0
   Exit Sub

cmdok_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure cmdok_Click of Formulario frmCE_Recepcion_Probetas"
End Sub

Private Sub cmdSalir_Click()
    Unload Me
End Sub
Private Sub Form_Load()
   On Error GoTo Form_Load_Error

    log (Me.Name)
    cargar_botones Me
    inicializar_grid
    If PK_MUESTRA > 0 Then
        cargar_recepcion
    End If
    'COC
    fecha = Date
    Dim oDeco As New clsDecodificadora
    oDeco.cargar_mi_combo cmbDimension(0), DECODIFICADORA.DECODIFICADORA_DIMENSIONES
    oDeco.cargar_mi_combo cmbDimension(1), DECODIFICADORA.DECODIFICADORA_DIMENSIONES
    cmbDimension(1).desactivar
    
    Dim NUMERO As String
    NUMERO = "01"
    Dim rs As ADODB.Recordset
    Set rs = datos_bd("select count(distinct id_muestra) from muestras where anulada = 0 and tipo_muestra_id = 294 and fecha_recepcion = current_date")
    If rs.RecordCount > 0 Then
        NUMERO = Format(rs(0), "00")
    End If
    txtCOC(0) = NUMERO
    txtCOC(2) = "0201"

   On Error GoTo 0
   Exit Sub

Form_Load_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure Form_Load of Formulario frmCE_Recepcion_Probetas"

End Sub
Private Sub inicializar_grid()
    x.ReDim 1, filas, 0, Col
    x.Clear
    Set grid.Array = x
    grid.Refresh
End Sub
Private Sub cargar_recepcion()
    Dim oCe_resultados As New clsCe_resultados
    Dim rs As ADODB.Recordset
   On Error GoTo cargar_recepcion_Error

    Set rs = oCe_resultados.Listado_por_muestra(PK_MUESTRA)
    Dim filas As Integer
    fila = 1
    If rs.RecordCount > 0 Then
        Do
            x(fila, COLS.DESIGNACION) = CStr(rs("DESIGNACION"))
            x(fila, COLS.PROBETA) = CStr(rs("PROBETA"))
            x(fila, COLS.AREA) = CStr(rs("AREA"))
            x(fila, COLS.IDEN_CLIENTE) = CStr(rs("IDENTIFICACION_CLIENTE"))
            x(fila, COLS.IDEN_CANAGROSA) = CStr(rs("IDENTIFICACION_CANAGROSA"))
            x(fila, COLS.dimension) = CStr(rs("DIMENSION"))
            x(fila, COLS.material) = CStr(rs("MATERIAL"))
            x(fila, COLS.CA) = CStr(rs("CRITERIO_ACEPTACION"))
            fila = fila + 1
            rs.MoveNext
        Loop Until rs.EOF
    End If
    contar_probetas
    grid.Refresh
    grid.Row = 0
    grid.Col = COLS.IDEN_CLIENTE
    Dim oMuestra As New clsMuestra
    If oMuestra.CargaMuestra(PK_MUESTRA) Then
        proteger_campos oMuestra.getCERRADA
    End If
    Dim oce_recepcion As New clsCe_recepcion
    If oce_recepcion.Carga(PK_MUESTRA) Then
        chkIL.Value = oce_recepcion.getIDENTIFICACION_LABORATORIO
    End If
   On Error GoTo 0
   Exit Sub

cargar_recepcion_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure cargar_recepcion of Formulario frmCE_Recepcion_Probetas"
    
End Sub
Private Sub contar_probetas()
    Dim i As Integer
    Dim CANTIDAD As Integer
    CANTIDAD = 0
    For i = 1 To filas
        If Not IsEmpty(x(i, COLS.IDEN_CANAGROSA)) Then
            If CStr(x(i, COLS.IDEN_CANAGROSA)) <> "" Then
                CANTIDAD = CANTIDAD + 1
            End If
        End If
    Next
    txtcantidad = CStr(CANTIDAD)
End Sub
Private Sub grid_KeyUp(KeyCode As Integer, Shift As Integer)
    contar_probetas
End Sub
Private Sub proteger_campos(CERRADA As Integer)
    
    ' Permiso para modificar la vida
    Dim permiso As Boolean
    permiso = False
    Dim op As New clsParametros
    Dim s() As String
    Dim i As Integer
    op.Carga parametros.PARAM_USUARIOS_MODIFICAN_EQUIPOS_MUESTRA_CERRADA, ""
    If op.getVALOR <> "" Then
        s = Split(op.getVALOR, ",")
        For i = LBound(s) To UBound(s)
            If USUARIO.getID_EMPLEADO = CInt(s(i)) Then
                permiso = True
            End If
        Next
    End If
    Set op = Nothing


    If CERRADA = 1 And permiso = False Then
        cmdok.Enabled = False
    Else
        cmdok.Enabled = True
    End If
    Select Case CERRADA
        Case 0
            lblCerrada = "ABIERTA"
        Case 1
            lblCerrada = "CERRADA"
        Case 2
            lblCerrada = "PTE. CIERRE"
        Case 3
            lblCerrada = "C.SIN INFORME"
    End Select
End Sub
Private Sub cmdInformarCOC_Click()
   On Error GoTo cmdInformarCOC_Click_Error

    If txtCOC(1) = "" Then
        MsgBox "Indique el Código, ejemplo 3124E", vbCritical, App.Title
        Exit Sub
    End If
    If cmbDimension(0).getTEXTO = "" Then
        MsgBox "Indique la dimension.", vbCritical, App.Title
        Exit Sub
    End If
    Dim i As Integer
    Dim f As String
    Dim Identificacion As String
    f = Format(fecha, "yymmdd")
    ' Numero de ensayo
    Dim NUMERO As String
    NUMERO = txtCOC(0)
    Dim letra As String
    letra = "A"
    Dim cadena As String
    cadena = txtCOC(2)
    
    Dim probetas As Integer
    probetas = 0
    For i = 1 To filas
        If Not IsEmpty(x(i, COLS.DESIGNACION)) Then
            Identificacion = f + NUMERO + "-" + txtCOC(1) + "-" + cadena + "-" + letra
            
            x(i, COLS.IDEN_CLIENTE) = Identificacion
            x(i, COLS.IDEN_CANAGROSA) = Identificacion
            x(i, COLS.dimension) = cmbDimension(0).getTEXTO
            letra = Chr(Asc(letra) + 1)
            probetas = probetas + 1
        End If
    Next
    If cmbDimension(1).getTEXTO <> "" Then
        For i = (probetas / 2) + 1 To filas
            If Not IsEmpty(x(i, COLS.DESIGNACION)) Then
                x(i, COLS.dimension) = cmbDimension(1).getTEXTO
            End If
        Next
    End If
    grid.Refresh
    frmCOC.visible = False
   On Error GoTo 0
   Exit Sub

cmdInformarCOC_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure cmdInformarCOC_Click of Formulario frmCE_Recepcion_Detalle_Probetas"
    
End Sub
Private Sub txtCOC_Change(Index As Integer)
    If Index = 1 Then
        If txtCOC(Index) <> "" Then
            If Mid(txtCOC(Index), 3, 1) = "2" Then
                cmbDimension(1).activar
                lblSet.visible = True
            Else
                cmbDimension(1).desactivar
                lblSet.visible = False
            End If
        End If
    End If
End Sub
