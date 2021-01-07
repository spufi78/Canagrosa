VERSION 5.00
Object = "{A49CE0E0-C0F9-11D2-B0EA-00A024695830}#1.0#0"; "tidate8.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{77DDF307-D82B-4757-8B3A-106EC9D75D4B}#8.0#0"; "tdbg8.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmProveedores_Detalle 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Mantenimiento de Proveedores"
   ClientHeight    =   10500
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   12660
   Icon            =   "frmProveedores_Detalle.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   10500
   ScaleWidth      =   12660
   StartUpPosition =   2  'CenterScreen
   Begin TDBDate6Ctl.TDBDate TDBDate1 
      Height          =   345
      Left            =   7830
      TabIndex        =   70
      Top             =   6255
      Visible         =   0   'False
      Width           =   2085
      _Version        =   65536
      _ExtentX        =   3678
      _ExtentY        =   609
      Calendar        =   "frmProveedores_Detalle.frx":030A
      Caption         =   "frmProveedores_Detalle.frx":0422
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmProveedores_Detalle.frx":048E
      Keys            =   "frmProveedores_Detalle.frx":04AC
      Spin            =   "frmProveedores_Detalle.frx":050A
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
   Begin VB.Frame frmBotones 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1005
      Left            =   0
      TabIndex        =   57
      Top             =   9405
      Width           =   12615
      Begin VB.CommandButton cmdCalidad 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Calidad"
         Height          =   870
         Left            =   4680
         Picture         =   "frmProveedores_Detalle.frx":0532
         Style           =   1  'Graphical
         TabIndex        =   75
         ToolTipText     =   "Enviar informe por E-mail"
         Top             =   45
         Width           =   1065
      End
      Begin VB.CommandButton cmdcancel 
         BackColor       =   &H00E0E0E0&
         Caption         =   "ESC-Salir"
         Height          =   870
         Left            =   11550
         Style           =   1  'Graphical
         TabIndex        =   27
         Top             =   45
         Width           =   1050
      End
      Begin VB.CommandButton cmdborrar 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Borrar Servicio Seleccionado"
         Height          =   930
         Left            =   90
         Picture         =   "frmProveedores_Detalle.frx":0DFC
         Style           =   1  'Graphical
         TabIndex        =   21
         Top             =   0
         Width           =   2280
      End
      Begin VB.CommandButton cmdEvaluacion 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Evaluación"
         Height          =   870
         Left            =   9255
         Picture         =   "frmProveedores_Detalle.frx":16C6
         Style           =   1  'Graphical
         TabIndex        =   25
         Top             =   45
         Width           =   1155
      End
      Begin VB.CommandButton cmdFacturas 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Facturas"
         Height          =   870
         Left            =   5790
         Picture         =   "frmProveedores_Detalle.frx":1F90
         Style           =   1  'Graphical
         TabIndex        =   23
         ToolTipText     =   "Enviar informe por E-mail"
         Top             =   45
         Width           =   1065
      End
      Begin VB.CommandButton cmdAdjuntar 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Adjuntos"
         Height          =   870
         Left            =   6885
         Style           =   1  'Graphical
         TabIndex        =   22
         ToolTipText     =   "Enviar informe por E-mail"
         Top             =   45
         Width           =   1155
      End
      Begin VB.CommandButton cmdRiesgo 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Riesgo"
         Height          =   870
         Left            =   8070
         Picture         =   "frmProveedores_Detalle.frx":285A
         Style           =   1  'Graphical
         TabIndex        =   24
         Top             =   45
         Width           =   1155
      End
      Begin VB.CommandButton cmdok 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Aceptar"
         Height          =   870
         Left            =   10455
         Style           =   1  'Graphical
         TabIndex        =   26
         Top             =   45
         Width           =   1050
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
         Height          =   375
         Left            =   4140
         TabIndex        =   58
         Top             =   225
         Visible         =   0   'False
         Width           =   4020
      End
   End
   Begin VB.Frame frmContabilidad 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Datos de Contabilidad y Facturación"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2625
      Left            =   45
      TabIndex        =   54
      Top             =   6750
      Width           =   12570
      Begin VB.Frame frmCuentaCobro 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Cuenta donde realizan el Cobro"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   780
         Left            =   5895
         TabIndex        =   67
         Top             =   1620
         Width           =   6450
         Begin MSDataListLib.DataCombo cmbBanco 
            Height          =   315
            Left            =   1035
            TabIndex        =   68
            Top             =   270
            Width           =   5325
            _ExtentX        =   9393
            _ExtentY        =   556
            _Version        =   393216
            Appearance      =   0
            Style           =   2
            Text            =   ""
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C0C0&
            Caption         =   "Banco"
            Height          =   195
            Index           =   16
            Left            =   90
            TabIndex        =   69
            Top             =   330
            Width           =   465
         End
      End
      Begin VB.Frame frmCuentaPago 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Cuenta de Pago"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   960
         Left            =   5895
         TabIndex        =   64
         Top             =   1620
         Width           =   6450
         Begin VB.TextBox txtdatos 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Index           =   10
            Left            =   1035
            MaxLength       =   50
            TabIndex        =   19
            Top             =   180
            Width           =   5370
         End
         Begin MSMask.MaskEdBox txtIBAN 
            Height          =   330
            Left            =   1035
            TabIndex        =   20
            Top             =   540
            Width           =   5370
            _ExtentX        =   9472
            _ExtentY        =   582
            _Version        =   393216
            Appearance      =   0
            MaxLength       =   29
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Mask            =   "&&##-####-####-####-####-####"
            PromptChar      =   "_"
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C0C0&
            Caption         =   "Banco"
            Height          =   195
            Index           =   12
            Left            =   90
            TabIndex        =   66
            Top             =   225
            Width           =   465
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C0C0&
            Caption         =   "IBAN"
            Height          =   195
            Index           =   14
            Left            =   90
            TabIndex        =   65
            Top             =   585
            Width           =   375
         End
      End
      Begin VB.Frame Frame4 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Datos de Pago"
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
         Left            =   5895
         TabIndex        =   59
         Top             =   225
         Width           =   6450
         Begin VB.TextBox txtdatos 
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
            Height          =   330
            Index           =   14
            Left            =   1035
            MaxLength       =   25
            TabIndex        =   17
            Top             =   945
            Width           =   1890
         End
         Begin VB.TextBox txtdatos 
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
            Height          =   330
            Index           =   9
            Left            =   4050
            MaxLength       =   25
            TabIndex        =   18
            Top             =   945
            Width           =   2295
         End
         Begin MSDataListLib.DataCombo cmbFP 
            Height          =   315
            Left            =   1035
            TabIndex        =   15
            Top             =   225
            Width           =   5325
            _ExtentX        =   9393
            _ExtentY        =   556
            _Version        =   393216
            Appearance      =   0
            Style           =   2
            Text            =   ""
         End
         Begin MSDataListLib.DataCombo cmbVencimiento 
            Height          =   315
            Left            =   1035
            TabIndex        =   16
            Top             =   585
            Width           =   5325
            _ExtentX        =   9393
            _ExtentY        =   556
            _Version        =   393216
            Appearance      =   0
            Style           =   2
            Text            =   ""
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C0C0&
            Caption         =   "Vencimiento"
            Height          =   195
            Index           =   19
            Left            =   90
            TabIndex        =   63
            Top             =   645
            Width           =   870
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C0C0&
            Caption         =   "Forma Pago"
            Height          =   195
            Index           =   8
            Left            =   90
            TabIndex        =   62
            Top             =   285
            Width           =   855
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C0C0&
            Caption         =   "C.Contable"
            Height          =   195
            Index           =   9
            Left            =   90
            TabIndex        =   61
            Top             =   1035
            Width           =   780
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C0C0&
            Caption         =   "C.Retención"
            Height          =   195
            Index           =   18
            Left            =   3060
            TabIndex        =   60
            Top             =   1035
            Width           =   885
         End
      End
      Begin VB.Frame Frame3 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Presupuesto y Gasto Real"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2355
         Index           =   0
         Left            =   45
         TabIndex        =   55
         Top             =   225
         Width           =   5820
         Begin TrueDBGrid80.TDBGrid gridPto 
            Height          =   1965
            Left            =   45
            TabIndex        =   56
            Top             =   225
            Width           =   5685
            _ExtentX        =   10028
            _ExtentY        =   3466
            _LayoutType     =   4
            _RowHeight      =   -2147483647
            _WasPersistedAsPixels=   0
            Columns(0)._VlistStyle=   0
            Columns(0)._MaxComboItems=   5
            Columns(0).Caption=   "Año"
            Columns(0).DataField=   ""
            Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(1)._VlistStyle=   0
            Columns(1)._MaxComboItems=   5
            Columns(1).Caption=   "Pto. Previsto"
            Columns(1).DataField=   ""
            Columns(1).NumberFormat=   "Currency"
            Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(2)._VlistStyle=   0
            Columns(2)._MaxComboItems=   5
            Columns(2).Caption=   "Gasto Real"
            Columns(2).DataField=   ""
            Columns(2).NumberFormat=   "Currency"
            Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns.Count   =   3
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
            Splits(0)._ColumnProps(0)=   "Columns.Count=3"
            Splits(0)._ColumnProps(1)=   "Column(0).Width=2646"
            Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
            Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=2540"
            Splits(0)._ColumnProps(4)=   "Column(0)._EditAlways=0"
            Splits(0)._ColumnProps(5)=   "Column(0)._ColStyle=1"
            Splits(0)._ColumnProps(6)=   "Column(0).Order=1"
            Splits(0)._ColumnProps(7)=   "Column(1).Width=3784"
            Splits(0)._ColumnProps(8)=   "Column(1).DividerColor=0"
            Splits(0)._ColumnProps(9)=   "Column(1)._WidthInPix=3678"
            Splits(0)._ColumnProps(10)=   "Column(1)._EditAlways=0"
            Splits(0)._ColumnProps(11)=   "Column(1)._ColStyle=2"
            Splits(0)._ColumnProps(12)=   "Column(1).Order=2"
            Splits(0)._ColumnProps(13)=   "Column(2).Width=900"
            Splits(0)._ColumnProps(14)=   "Column(2).DividerColor=0"
            Splits(0)._ColumnProps(15)=   "Column(2)._WidthInPix=794"
            Splits(0)._ColumnProps(16)=   "Column(2)._EditAlways=0"
            Splits(0)._ColumnProps(17)=   "Column(2)._ColStyle=8194"
            Splits(0)._ColumnProps(18)=   "Column(2).AllowFocus=0"
            Splits(0)._ColumnProps(19)=   "Column(2).Order=3"
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
            _StyleDefs(6)   =   "Style:id=1,.parent=0,.namedParent=37,.bgcolor=&HDEEDFA&,.bold=0,.fontsize=975"
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
            _StyleDefs(40)  =   "Splits(0).Columns(1).Style:id=62,.parent=11,.alignment=1"
            _StyleDefs(41)  =   "Splits(0).Columns(1).HeadingStyle:id=59,.parent=12"
            _StyleDefs(42)  =   "Splits(0).Columns(1).FooterStyle:id=60,.parent=13"
            _StyleDefs(43)  =   "Splits(0).Columns(1).EditorStyle:id=61,.parent=15"
            _StyleDefs(44)  =   "Splits(0).Columns(2).Style:id=32,.parent=11,.alignment=1,.bgcolor=&HC0C0C0&"
            _StyleDefs(45)  =   ":id=32,.locked=-1"
            _StyleDefs(46)  =   "Splits(0).Columns(2).HeadingStyle:id=29,.parent=12"
            _StyleDefs(47)  =   "Splits(0).Columns(2).FooterStyle:id=30,.parent=13"
            _StyleDefs(48)  =   "Splits(0).Columns(2).EditorStyle:id=31,.parent=15"
            _StyleDefs(49)  =   "Named:id=37:Normal"
            _StyleDefs(50)  =   ":id=37,.parent=0,.alignment=3,.bgcolor=&H80000018&,.bgpicMode=2,.bold=0"
            _StyleDefs(51)  =   ":id=37,.fontsize=825,.italic=0,.underline=0,.strikethrough=0,.charset=0"
            _StyleDefs(52)  =   ":id=37,.fontname=MS Sans Serif"
            _StyleDefs(53)  =   "Named:id=38:Heading"
            _StyleDefs(54)  =   ":id=38,.parent=37,.valignment=2,.bgcolor=&H80000016&,.fgcolor=&H80000012&"
            _StyleDefs(55)  =   ":id=38,.wraptext=-1,.bold=0,.fontsize=825,.italic=0,.underline=0"
            _StyleDefs(56)  =   ":id=38,.strikethrough=0,.charset=0"
            _StyleDefs(57)  =   ":id=38,.fontname=MS Sans Serif"
            _StyleDefs(58)  =   "Named:id=39:Footing"
            _StyleDefs(59)  =   ":id=39,.parent=37,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
            _StyleDefs(60)  =   "Named:id=40:Selected"
            _StyleDefs(61)  =   ":id=40,.parent=37,.bgcolor=&H8080FF&,.fgcolor=&H0&,.bold=0,.fontsize=825"
            _StyleDefs(62)  =   ":id=40,.italic=0,.underline=0,.strikethrough=0,.charset=0"
            _StyleDefs(63)  =   ":id=40,.fontname=MS Sans Serif"
            _StyleDefs(64)  =   "Named:id=41:Caption"
            _StyleDefs(65)  =   ":id=41,.parent=38,.alignment=2"
            _StyleDefs(66)  =   "Named:id=42:HighlightRow"
            _StyleDefs(67)  =   ":id=42,.parent=37,.bgcolor=&H80000008&,.fgcolor=&H80000005&"
            _StyleDefs(68)  =   "Named:id=43:EvenRow"
            _StyleDefs(69)  =   ":id=43,.parent=37,.bgcolor=&HFFFFFF&,.wraptext=-1"
            _StyleDefs(70)  =   "Named:id=44:OddRow"
            _StyleDefs(71)  =   ":id=44,.parent=37,.bgcolor=&HD5ECF9&"
            _StyleDefs(72)  =   "Named:id=47:RecordSelector"
            _StyleDefs(73)  =   ":id=47,.parent=38"
            _StyleDefs(74)  =   "Named:id=50:FilterBar"
            _StyleDefs(75)  =   ":id=50,.parent=37"
         End
      End
   End
   Begin VB.Frame Frame5 
      Caption         =   "Frame5"
      Height          =   15
      Left            =   3690
      TabIndex        =   53
      Top             =   11205
      Width           =   645
   End
   Begin TrueDBGrid80.TDBDropDown tMetodos 
      Height          =   2280
      Left            =   4860
      TabIndex        =   48
      Top             =   5355
      Width           =   2235
      _ExtentX        =   3942
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
      Splits(0)._ColumnProps(1)=   "Column(0).Width=2725"
      Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
      Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=2646"
      Splits(0)._ColumnProps(4)=   "Column(0)._EditAlways=0"
      Splits(0)._ColumnProps(5)=   "Column(0).Order=1"
      Splits(0)._ColumnProps(6)=   "Column(1).Width=2725"
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
   Begin VB.Frame Frame2 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Observaciones "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1155
      Index           =   13
      Left            =   45
      TabIndex        =   39
      Top             =   3780
      Width           =   12570
      Begin VB.TextBox txtdatos 
         Appearance      =   0  'Flat
         Height          =   870
         Index           =   13
         Left            =   135
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   14
         Top             =   225
         Width           =   12315
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0C0C0&
      Caption         =   " Datos del Proveedor"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3180
      Left            =   45
      TabIndex        =   28
      Top             =   540
      Width           =   12570
      Begin VB.CheckBox chkAnulado 
         BackColor       =   &H00C0C0C0&
         Caption         =   "ANULADO"
         ForeColor       =   &H000000FF&
         Height          =   240
         Left            =   11340
         TabIndex        =   74
         Top             =   2790
         Width           =   1140
      End
      Begin VB.TextBox txtdatos 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   330
         Index           =   11
         Left            =   6450
         TabIndex        =   13
         Text            =   "3"
         Top             =   2745
         Width           =   960
      End
      Begin VB.CheckBox chkExtra 
         BackColor       =   &H00C0C0C0&
         Caption         =   "ExtraComunitario"
         Height          =   240
         Left            =   1935
         TabIndex        =   72
         Top             =   2835
         Width           =   1860
      End
      Begin VB.CheckBox chkIntra 
         BackColor       =   &H00C0C0C0&
         Caption         =   "IntraComunitario"
         Height          =   240
         Left            =   135
         TabIndex        =   71
         Top             =   2835
         Width           =   1680
      End
      Begin VB.TextBox txtdatos 
         Appearance      =   0  'Flat
         Height          =   330
         Index           =   0
         Left            =   1080
         TabIndex        =   9
         Top             =   2025
         Width           =   4245
      End
      Begin VB.CheckBox chkFormador 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Formador"
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
         Left            =   4140
         TabIndex        =   51
         Top             =   1350
         Width           =   1140
      End
      Begin VB.TextBox txtdatos 
         Appearance      =   0  'Flat
         Height          =   330
         Index           =   8
         Left            =   1080
         MaxLength       =   25
         TabIndex        =   11
         Top             =   2385
         Width           =   4245
      End
      Begin VB.TextBox txtdatos 
         Appearance      =   0  'Flat
         Height          =   330
         Index           =   12
         Left            =   6450
         MaxLength       =   25
         TabIndex        =   12
         Top             =   2385
         Width           =   5955
      End
      Begin VB.CheckBox chkSubcontrata 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Subcontrata"
         ForeColor       =   &H00FF0000&
         Height          =   240
         Left            =   11160
         TabIndex        =   40
         Top             =   0
         Width           =   1275
      End
      Begin MSDataListLib.DataCombo cmbPais 
         Height          =   315
         Left            =   6450
         TabIndex        =   4
         Top             =   960
         Width           =   5970
         _ExtentX        =   10530
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         Text            =   ""
      End
      Begin VB.TextBox txtdatos 
         Appearance      =   0  'Flat
         Height          =   330
         Index           =   5
         Left            =   1080
         MaxLength       =   30
         TabIndex        =   7
         Top             =   1665
         Width           =   2460
      End
      Begin VB.TextBox txtdatos 
         Appearance      =   0  'Flat
         Height          =   330
         Index           =   4
         Left            =   1080
         MaxLength       =   30
         TabIndex        =   5
         Top             =   1305
         Width           =   2460
      End
      Begin VB.TextBox txtdatos 
         Appearance      =   0  'Flat
         Height          =   330
         Index           =   7
         Left            =   6450
         TabIndex        =   10
         Top             =   2025
         Width           =   5955
      End
      Begin VB.TextBox txtdatos 
         Appearance      =   0  'Flat
         Height          =   330
         Index           =   6
         Left            =   1080
         MaxLength       =   15
         TabIndex        =   2
         Top             =   960
         Width           =   2460
      End
      Begin VB.TextBox txtdatos 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   330
         Index           =   3
         Left            =   4275
         TabIndex        =   3
         Top             =   960
         Width           =   960
      End
      Begin VB.TextBox txtdatos 
         Appearance      =   0  'Flat
         Height          =   330
         Index           =   2
         Left            =   1080
         MaxLength       =   75
         TabIndex        =   1
         Top             =   600
         Width           =   11355
      End
      Begin VB.TextBox txtdatos 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   330
         Index           =   1
         Left            =   1080
         MaxLength       =   75
         TabIndex        =   0
         Top             =   240
         Width           =   11355
      End
      Begin MSDataListLib.DataCombo cmbProvincia 
         Height          =   315
         Left            =   6450
         TabIndex        =   6
         Top             =   1335
         Width           =   5970
         _ExtentX        =   10530
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         Text            =   ""
      End
      Begin MSDataListLib.DataCombo cmbMunicipios 
         Height          =   315
         Left            =   6450
         TabIndex        =   8
         Top             =   1695
         Width           =   5970
         _ExtentX        =   10530
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         Text            =   ""
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Días Entrega"
         Height          =   195
         Index           =   20
         Left            =   5445
         TabIndex        =   73
         Top             =   2820
         Width           =   945
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "E-mail (Fac.)"
         Height          =   195
         Index           =   17
         Left            =   135
         TabIndex        =   52
         Top             =   2115
         Width           =   870
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Contacto"
         Height          =   195
         Index           =   13
         Left            =   135
         TabIndex        =   50
         Top             =   2475
         Width           =   645
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Pagina Web"
         Height          =   195
         Index           =   10
         Left            =   5445
         TabIndex        =   49
         Top             =   2460
         Width           =   885
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Municipio"
         Height          =   195
         Index           =   7
         Left            =   5460
         TabIndex        =   38
         Top             =   1755
         Width           =   675
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "FAX"
         Height          =   195
         Index           =   15
         Left            =   150
         TabIndex        =   37
         Top             =   1725
         Width           =   300
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Telefono"
         Height          =   195
         Index           =   11
         Left            =   135
         TabIndex        =   36
         Top             =   1365
         Width           =   630
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "E-mail"
         Height          =   195
         Index           =   6
         Left            =   5445
         TabIndex        =   35
         Top             =   2115
         Width           =   420
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "C.I.F."
         Height          =   195
         Index           =   5
         Left            =   135
         TabIndex        =   34
         Top             =   1035
         Width           =   375
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Provincia"
         Height          =   195
         Index           =   4
         Left            =   5445
         TabIndex        =   33
         Top             =   1395
         Width           =   660
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Pais"
         Height          =   195
         Index           =   3
         Left            =   5460
         TabIndex        =   32
         Top             =   1020
         Width           =   300
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "C.P."
         Height          =   195
         Index           =   2
         Left            =   3870
         TabIndex        =   31
         Top             =   1035
         Width           =   300
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Direccion"
         Height          =   195
         Index           =   1
         Left            =   135
         TabIndex        =   30
         Top             =   645
         Width           =   675
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Nombre"
         Height          =   195
         Index           =   0
         Left            =   135
         TabIndex        =   29
         Top             =   285
         Width           =   555
      End
   End
   Begin TrueDBGrid80.TDBDropDown tServicios 
      Height          =   2280
      Left            =   90
      TabIndex        =   46
      Top             =   5400
      Width           =   2595
      _ExtentX        =   4577
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
      Splits(0)._ColumnProps(1)=   "Column(0).Width=2725"
      Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
      Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=2646"
      Splits(0)._ColumnProps(4)=   "Column(0)._EditAlways=0"
      Splits(0)._ColumnProps(5)=   "Column(0).Order=1"
      Splits(0)._ColumnProps(6)=   "Column(1).Width=2725"
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
   Begin TrueDBGrid80.TDBGrid grid 
      Height          =   1695
      Left            =   45
      TabIndex        =   47
      Top             =   4995
      Width           =   12525
      _ExtentX        =   22093
      _ExtentY        =   2990
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
      Splits(0)._ColumnProps(1)=   "Column(0).Width=4551"
      Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
      Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=4471"
      Splits(0)._ColumnProps(4)=   "Column(0)._EditAlways=0"
      Splits(0)._ColumnProps(5)=   "Column(0)._ColStyle=1"
      Splits(0)._ColumnProps(6)=   "Column(0).Order=1"
      Splits(0)._ColumnProps(7)=   "Column(1).Width=3916"
      Splits(0)._ColumnProps(8)=   "Column(1).DividerColor=0"
      Splits(0)._ColumnProps(9)=   "Column(1)._WidthInPix=3836"
      Splits(0)._ColumnProps(10)=   "Column(1)._EditAlways=0"
      Splits(0)._ColumnProps(11)=   "Column(1).Order=2"
      Splits(0)._ColumnProps(12)=   "Column(2).Width=3916"
      Splits(0)._ColumnProps(13)=   "Column(2).DividerColor=0"
      Splits(0)._ColumnProps(14)=   "Column(2)._WidthInPix=3836"
      Splits(0)._ColumnProps(15)=   "Column(2)._EditAlways=0"
      Splits(0)._ColumnProps(16)=   "Column(2)._ColStyle=1"
      Splits(0)._ColumnProps(17)=   "Column(2).Order=3"
      Splits(0)._ColumnProps(18)=   "Column(3).Width=2064"
      Splits(0)._ColumnProps(19)=   "Column(3).DividerColor=0"
      Splits(0)._ColumnProps(20)=   "Column(3)._WidthInPix=1984"
      Splits(0)._ColumnProps(21)=   "Column(3)._EditAlways=0"
      Splits(0)._ColumnProps(22)=   "Column(3)._ColStyle=1"
      Splits(0)._ColumnProps(23)=   "Column(3).Order=4"
      Splits(0)._ColumnProps(24)=   "Column(4).Width=2196"
      Splits(0)._ColumnProps(25)=   "Column(4).DividerColor=0"
      Splits(0)._ColumnProps(26)=   "Column(4)._WidthInPix=2117"
      Splits(0)._ColumnProps(27)=   "Column(4)._EditAlways=0"
      Splits(0)._ColumnProps(28)=   "Column(4)._ColStyle=1"
      Splits(0)._ColumnProps(29)=   "Column(4).Order=5"
      Splits(0)._ColumnProps(30)=   "Column(4).AutoDropDown=1"
      Splits(0)._ColumnProps(31)=   "Column(5).Width=1773"
      Splits(0)._ColumnProps(32)=   "Column(5).DividerColor=0"
      Splits(0)._ColumnProps(33)=   "Column(5)._WidthInPix=1693"
      Splits(0)._ColumnProps(34)=   "Column(5)._EditAlways=0"
      Splits(0)._ColumnProps(35)=   "Column(5).Order=6"
      Splits(0)._ColumnProps(36)=   "Column(6).Width=2725"
      Splits(0)._ColumnProps(37)=   "Column(6).DividerColor=0"
      Splits(0)._ColumnProps(38)=   "Column(6)._WidthInPix=2646"
      Splits(0)._ColumnProps(39)=   "Column(6)._EditAlways=0"
      Splits(0)._ColumnProps(40)=   "Column(6).Visible=0"
      Splits(0)._ColumnProps(41)=   "Column(6).Order=7"
      Splits(0)._ColumnProps(42)=   "Column(7).Width=2725"
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
   Begin VB.Frame Frame3 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Adjuntos"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1770
      Index           =   3
      Left            =   4230
      TabIndex        =   42
      Top             =   45
      Visible         =   0   'False
      Width           =   6675
      Begin VB.CommandButton cmdEliminarAdjunto 
         BackColor       =   &H00E0E0E0&
         Height          =   465
         Index           =   3
         Left            =   6120
         Picture         =   "frmProveedores_Detalle.frx":3124
         Style           =   1  'Graphical
         TabIndex        =   44
         ToolTipText     =   "Eliminar accesorio"
         Top             =   1215
         Width           =   465
      End
      Begin VB.CommandButton cmdEscaner 
         BackColor       =   &H00E0E0E0&
         Height          =   465
         Left            =   6120
         Picture         =   "frmProveedores_Detalle.frx":32B8
         Style           =   1  'Graphical
         TabIndex        =   43
         Top             =   495
         Visible         =   0   'False
         Width           =   465
      End
      Begin MSComctlLib.ListView lista 
         Height          =   1470
         Left            =   135
         TabIndex        =   45
         Top             =   225
         Width           =   5940
         _ExtentX        =   10478
         _ExtentY        =   2593
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
      TabIndex        =   41
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
      Width           =   12700
   End
End
Attribute VB_Name = "frmProveedores_Detalle"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public PK As Long

Dim X As New XArrayDB

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


Dim xPto As New XArrayDB
Const filasPto As Integer = 10
Const ColPto As Integer = 2
Private Enum COLSPTO
    ANO = 0
    PRESUPUESTO = 1
    REAL = 2
End Enum

Private Sub cmbFP_Change()
    Dim oDeco As New clsDecodificadora
    frmCuentaCobro.visible = False
    frmCuentaPago.visible = False
    cmbVencimiento.Enabled = False
    cmbVencimiento.BoundText = "0"
    If cmbFP.Text <> "" Then
        oDeco.Carga_valor DECODIFICADORA.DECODIFICADORA_PROVEEDORES_FP, cmbFP.BoundText
        Dim s() As String
        Dim s2() As String
        s = Split(oDeco.getPARAMETROS, ";")
        ' Vencimiento
        s2 = Split(s(0), "=")
        If s2(1) = "1" Then
            cmbVencimiento.Enabled = True
        End If
        ' Banco
        s2 = Split(s(1), "=")
        If s2(1) = "0" Then
            frmCuentaPago.visible = True
        Else
            frmCuentaCobro.visible = True
        End If
    End If
    Set oDeco = Nothing
End Sub

Private Sub cmdAdjuntar_Click()
    If PK > 0 Then
        With frmAdjuntos
            .TOBJETO = TOBJETO.TOBJETO_PROVEEDOR
            .COBJETO = PK
            .Show 1
        End With
        Set frmAdjuntos = Nothing
    End If
End Sub
Private Sub cmdAdjuntos_Click()
End Sub

Private Sub cmdBorrar_Click()
    On Error Resume Next
    Dim i As Integer
    For i = 0 To Col
        grid.SelBookmarks.Add grid.Bookmark
        X(grid.Bookmark, i) = ""
        grid.SelBookmarks.Remove 0
    Next
    grid.Refresh
    grid.SetFocus
End Sub


Private Sub cmbPais_LostFocus()
    cargar_provincias
End Sub
Private Sub cmbProvincia_LostFocus()
    cargar_municipios
End Sub

Private Sub cmdCalidad_Click()
    If PK > 0 Then
        frmProveedores_Calidad.PK = PK
        frmProveedores_Calidad.Show 1
    End If
End Sub

Private Sub cmdcancel_Click()
    Unload Me
End Sub

Private Sub cmdEliminar_Click(Index As Integer)
    If lista.ListItems.Count > 0 Then
        lista.ListItems.Remove lista.selectedItem.Index
    End If

End Sub

Private Sub cmdEliminarAdjunto_Click(Index As Integer)
    If lista.ListItems.Count > 0 Then
        lista.ListItems.Remove lista.selectedItem.Index
    End If

End Sub

Private Sub cmdEscaner_Click()
   On Error GoTo cmdEscaner_Click_Error

    If PK = 0 Then
        Dim c As String
        
        c = "Para añadir adjuntos, es necesario primero insertar el proveedor."
        c = c & vbNewLine & " Pulse aceptar, para insertar al proveedor en el sistema y "
        c = c & vbNewLine & " posteriormente añada los adjuntos que desee. "
        
        MsgBox c, vbInformation, App.Title
        Exit Sub
    End If
        
    frmEscaner.Show 1
    If documento_escaner <> "" Then
        Dim nombreNuevo As String
        nombreNuevo = ""
        nombreNuevo = InputBox("Introduzca nombre para el archivo Escaneado, sin ninguna extensión (SOLO LETRAS Y NUMEROS).", "Escaneando Archivo", , Screen.Width / 3, (Screen.Height / 3))
        nombreNuevo = Eliminar_Caracteres_Archivo(nombreNuevo)
        If Trim(nombreNuevo) <> "" Then
            If Dir(documento_escaner) = "" Then
                MsgBox "El documento que intenta vincular no existe en la ruta.", vbExclamation, App.Title
                Exit Sub
            End If
            On Error Resume Next
            Dim RUTA As String
            RUTA = ReadINI(App.Path + "\config.ini", "Documentos", "proveedores_adjuntos")
            MkDir RUTA & "\" & CStr(PK)
            FileCopy documento_escaner, RUTA & "\" & CStr(PK) & "\" & nombreNuevo & ".pdf"
            With lista.ListItems.Add(, , lista.ListItems.Count + 1)
                .SubItems(1) = nombreNuevo
                .SubItems(2) = nombreNuevo & ".pdf"
            End With
            MsgBox "Fichero vinculado correctamente.", vbInformation, App.Title
        End If
    End If

   On Error GoTo 0
   Exit Sub

cmdEscaner_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure cmdEscaner_Click of Formulario frmEmpleados_Cualificaciones_Nueva"

End Sub

Private Sub cmdEvaluacion_Click()
    frmProveedores_Evaluacion.PK = PK
    frmProveedores_Evaluacion.Show 1

End Sub

Private Sub cmdFacturas_Click()
    If PK > 0 Then
'M1257-I
        frmProveedores_Facturas.TOBJETO = 0
        frmProveedores_Facturas.COBJETO = 0
'M1257-F
        frmProveedores_Facturas.PK = PK
        frmProveedores_Facturas.Show 1
    End If

End Sub

Private Sub cmdok_Click()
    'E0071-I
    'If gproveedor <> 0 Then
    If PK <> 0 Then
    'E0071-F
        modificar_proveedor
    Else
        insertar_proveedor
    End If
End Sub

Private Sub cmdRiesgo_Click()
    frmProveedores_Riesgo.PK = PK
    frmProveedores_Riesgo.Show 1
End Sub

Private Sub Form_Load()
    log (Me.Name)
    cargar_botones Me
    cargar_paises
    cabecera
    inicializar_grid
    cargar_combos
    'E0072-I
    'If gproveedor <> 0 Then
    If PK <> 0 Then
    'E0072-F
        consulta_proveedor
    Else
        cmdFacturas.visible = False
        cmdAdjuntar.visible = False
        cmdEvaluacion.visible = False
        cmdRiesgo.visible = False
        cmbFP.BoundText = 0
        'TESORERIA-I
        Dim oParametro As New clsParametros
        oParametro.Carga parametros.PARAM_CUENTA_CONTABLE_PROVEEDOR, ""
        txtdatos(14) = oParametro.getVALOR
        Set oParametro = Nothing
        'TESORERIA-F
    End If
    numero_servicios
    If USUARIO.getPER_TESORERIA_MENU = False Then
        cmdborrar.visible = False
        cmdok.visible = False
    End If
    If USUARIO.getPER_TESORERIA_FP = False Then
        cmdFacturas.visible = False
        frmContabilidad.visible = False
        frmBotones.top = frmContabilidad.top
        Me.Height = Me.Height - frmContabilidad.Height
    End If

End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set frmProveedores_Detalle = Nothing
End Sub

Private Sub grid_AfterColEdit(ByVal ColIndex As Integer)
   On Error GoTo grid_AfterColEdit_Error
    numero_servicios
'    Select Case ColIndex
'       Case COLS.finicial
'        grid.Row = grid.Row + 1
'       Case COLS.fAprobacion
'        grid.Row = grid.Row + 1
'       Case COLS.OBSERVACIONES
'        grid.Row = grid.Row + 1
'    End Select
'    grid.Refresh
   On Error GoTo 0
   Exit Sub

grid_AfterColEdit_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure grid_AfterColEdit of Formulario frmDocumento_Edicion"

End Sub

Private Sub lista_DblClick()
    If lista.ListItems.Count = 0 Then
        Exit Sub
    End If
    On Error GoTo fallo
    Dim RUTA As String
    RUTA = ReadINI(App.Path + "\config.ini", "Documentos", "proveedores_adjuntos")
    RUTA = RUTA & "\" & CStr(PK)
    RUTA = RUTA & "\" & lista.ListItems(lista.selectedItem.Index).SubItems(2)
    If RUTA <> "" Then
        If Dir(RUTA) <> "" Then
            Dim r As Long
            r = Shell("rundll32.exe url.dll,FileProtocolHandler " & RUTA, vbMaximizedFocus)
        Else
            MsgBox "El documento vinculado no existe.", vbCritical, App.Title
        End If
    End If
    Exit Sub
fallo:
    MsgBox "Error al abrir el documento.", vbCritical, App.Title
End Sub

Private Sub txtdatos_GotFocus(Index As Integer)
    txtdatos(Index).BackColor = &H80C0FF
    txtdatos(Index).SelStart = 0
    txtdatos(Index).SelLength = Len(txtdatos(Index))
End Sub

Private Sub txtdatos_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
     Case 40
       If Index = 15 Then
        txtdatos(1).SetFocus
       Else
        SendKeys "{Tab}", True
       End If
       'E0066-I
       ' se comenta porque no encuentra la variable (ya está a 0 en el keypress)
       'KeyAscii = 0 ' Para evitar el "bip" del sistema
       'E0066-F
     Case 38
       If Index = 1 Then
        txtdatos(15).SetFocus
       Else
        SendKeys "+{Tab}", True
       End If
       'E0067-I
       'KeyAscii = 0 ' Para evitar el "bip" del sistema
       'E0067-F
     Case 27
        cmdcancel_Click
    End Select
End Sub

Private Sub txtDatos_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = 13 And Index <> 16 Then
       If Index = 15 Then
        txtdatos(1).SetFocus
        KeyAscii = 0 ' Para evitar el "bip" del sistema
       Else
        SendKeys "{Tab}", True
        KeyAscii = 0 ' Para evitar el "bip" del sistema
       End If
    End If
End Sub

Private Sub txtdatos_LostFocus(Index As Integer)
    txtdatos(Index).BackColor = &HFFFFFF
End Sub

Public Sub borrar_campos()
    Dim i As Integer
    For i = 1 To 13
       If i < 9 Or i > 11 Then
        txtdatos(i) = ""
       End If
    Next
    cmbPais.Text = ""
    cmbProvincia.Text = ""
    cmbMunicipios.Text = ""
    txtdatos(1).SetFocus
End Sub

Public Sub bloquear_campos()
    Dim i As Integer
    For i = 1 To 13
       If i < 9 Or i > 11 Then
        txtdatos(i).Locked = True
       End If
    Next
    cmbMunicipios.Locked = True
    cmbProvincia.Locked = True
    cmbPais.Locked = True
End Sub

Public Sub insertar_proveedor()
    If valida_datos = False Then
        Exit Sub
    End If
    If MsgBox("Va a dar de alta el proveedor. ¿Esta seguro?", vbQuestion + vbYesNo, "Informacion") = vbYes Then
        'E0068-I
        'Se declara la variable porque el set de abajo no la encuentra
        Dim oProveedor As New clsProveedor
        'E0068-F
        Set oProveedor = mover_datos
        PK = oProveedor.Insertar
        If (PK > 0) Then
            insertar_servicios
            insertarPto
            If MsgBox("Proveedor almacenado correctamente. ¿Desea añadir adjuntos?", vbQuestion + vbYesNo, App.Title) = vbNo Then
                Unload Me
            End If
        End If
'        borrar_campos
        Set oProveedor = Nothing
    End If
End Sub

Public Sub modificar_proveedor()
    If valida_datos() = False Then
        Exit Sub
    End If
    Dim pos As Integer
    Dim proveedor As Integer
    If MsgBox("Va a modificar los datos del proveedor. ¿Esta seguro?", vbQuestion + vbYesNo, "Informacion") = vbYes Then
        'E0069-I
        'Se declara la variable porque el set de abajo no la encuentra
        Dim oProveedor As New clsProveedor
        'E0069-F
        Set oProveedor = mover_datos
        'E0073-I
        'oProveedor.setID_PROVEEDOR = gproveedor
        oProveedor.setID_PROVEEDOR = PK
        'E0073-F
        If oProveedor.Modificar = True Then
            insertar_adjuntos
            insertar_servicios
            insertarPto
            Unload Me
        End If
        Set oProveedor = Nothing
    End If

End Sub


Public Function valida_datos() As Boolean
    valida_datos = True
    If txtdatos(1) = "" Then
        MsgBox "El nombre del proveedor no puede estar en blanco.", vbCritical, "Error"
        txtdatos(1).SetFocus
        valida_datos = False
        Exit Function
    End If
    If txtdatos(2) = "" Then
        MsgBox "La dirección no puede estar en blanco.", vbCritical, "Error"
        txtdatos(2).SetFocus
        valida_datos = False
        Exit Function
    End If
    If txtdatos(3) = "" Then
        MsgBox "El codigo postal no puede estar en blanco.", vbCritical, "Error"
        txtdatos(3).SetFocus
        valida_datos = False
        Exit Function
    Else
        If Not IsNumeric(txtdatos(3)) Then
            MsgBox "El codigo postal debe ser numérico.", vbCritical, "Error"
            txtdatos(3).SetFocus
            valida_datos = False
            Exit Function
        End If
    End If
    If txtdatos(6) = "" Then
        MsgBox "El CIF no puede estar en blanco.", vbCritical, "Error"
        txtdatos(6).SetFocus
        valida_datos = False
        Exit Function
    End If
    If cmbPais.Text = "" Then
        MsgBox "El pais no puede estar en blanco.", vbCritical, "Error"
        cmbPais.SetFocus
        valida_datos = False
        Exit Function
    End If
    If cmbProvincia.Text = "" Then
        MsgBox "La provincia no puede estar en blanco.", vbCritical, "Error"
        cmbProvincia.SetFocus
        valida_datos = False
        Exit Function
    End If
    If cmbMunicipios.Text = "" Then
        MsgBox "El municipio no puede estar en blanco.", vbCritical, "Error"
        cmbMunicipios.SetFocus
        valida_datos = False
        Exit Function
    End If
    'TESORERIA-I
    If txtdatos(14) = "" Then
        MsgBox "La cuenta contable no puede estar en blanco.", vbCritical, "Error"
        txtdatos(14).SetFocus
        valida_datos = False
        Exit Function
    End If
    If Len(txtdatos(14)) <> 7 Then
        MsgBox "La cuenta contable debe ser de 7 numeros.", vbCritical, "Error"
        txtdatos(14).SetFocus
        valida_datos = False
        Exit Function
    End If
    If Len(txtdatos(9)) > 0 And Len(txtdatos(9)) <> 7 Then
        MsgBox "La cuenta de retenciones debe ser de 7 numeros o dejarla vacía.", vbCritical, "Error"
        txtdatos(9).SetFocus
        valida_datos = False
        Exit Function
    End If
    'TESORERIA-F
    ' Validar servicios
    Dim i As Integer
    For i = 0 To filas
        If Trim(X(i, COLS.servicio)) <> "" Then
            If X(i, COLS.METODO) = "" Then
                MsgBox "Introduzca el metodo para los servicios añadidos.", vbExclamation, App.Title
                valida_datos = False
                Exit Function
            End If
            If X(i, COLS.finicial) = "" Then
                MsgBox "Introduzca la fecha inicial para los servicios.", vbExclamation, App.Title
                valida_datos = False
                Exit Function
            End If
        End If
    Next
    ' Validar Presupuesto
    For i = 0 To filasPto
        If Trim(xPto(i, COLSPTO.ANO)) <> "" Then
            If Not IsNumeric(xPto(i, COLSPTO.ANO)) Then
                MsgBox "Introduzca los años del presupuesto correctamente.", vbExclamation, App.Title
                valida_datos = False
                Exit Function
            End If
        End If
    Next
    ' VALIDAR IBAN
    If txtIBAN.Text <> "" Then
        Dim pais As String
        Dim iban As String
        Dim ibanCalculado As String
        pais = Left(txtIBAN.Text, 2)
        iban = Left(txtIBAN.Text, 4)
        If pais <> "__" Then
            ibanCalculado = Left(CalcularIBAN(pais, Right(txtIBAN.Text, Len(txtIBAN.Text) - 5)), 4)
            If iban <> ibanCalculado Then
                    MsgBox "El IBAN introducido no es correcto.", vbExclamation, App.Title
                    valida_datos = False
                    Exit Function
            End If
        End If
    End If
    If txtdatos(11) = "" Then
        MsgBox "Los días de entrega no pueden estar en blanco.", vbCritical, "Error"
        txtdatos(11).SetFocus
        valida_datos = False
        Exit Function
    End If
    
End Function

Public Sub consulta_proveedor()
    On Error GoTo fallo
    Dim oProveedor As New clsProveedor
    Dim strCuenta As String
    lbltitulo.Caption = "MODIFICACIÓN DE PROVEEDOR"
'    lbltitulo.BackColor = &H80C0FF
    'E0073-I
    'oProveedor.Carga (gproveedor)
    oProveedor.Carga (PK)
    'E0073-F
    With oProveedor
        'M1339-I
        If Not IsNull(.getEMAIL_FACTURACION) Then
            txtdatos(0) = .getEMAIL_FACTURACION
        Else
            txtdatos(0) = ""
        End If
        'M1339-F
        txtdatos(1) = .getNOMBRE
        txtdatos(2) = .getDIRECCION
        txtdatos(3) = .getCOD_POSTAL
        txtdatos(6) = .getCIF
        txtdatos(4) = .getTELEFONO
        txtdatos(5) = .getFAX
        txtdatos(8) = .getRESPONSABLE
        txtdatos(7) = .getEMAIL
        txtdatos(13) = .getOBSERVACIONES
        'TESORERIA-I
        txtdatos(14) = .getCC
        txtdatos(9) = .getCC_RETENCION
        'TESORERIA-F
        txtdatos(12) = .getWEB
        'E0200-I
        txtdatos(11) = .getDIAS_ENTREGA
        If .getES_SUBCONTRATA = 0 Then
            chkSubcontrata.Value = Unchecked
        Else
            chkSubcontrata.Value = Checked
        End If
        'E0200-F
        'MXXXX-I
        If .getES_FORMADOR = 0 Then
            chkFormador.Value = Unchecked
        Else
            chkFormador.Value = Checked
        End If
        chkIntra = .getINTRA
        chkExtra = .getEXTRA
        'MXXXX-F
        'M1334-I
'        If .getES_CUENTA_ESPANYA = 0 Then
'            chkIBAN.value = Unchecked
'        Else
'            chkIBAN.value = Checked
'        End If
'        strCuenta = Trim(.getCUENTA_BANCARIA)
'        txtCuentaIBAN.Text = strCuenta
'        chkIBAN_Click
        txtdatos(10) = .getBANCO
        If .getCUENTA_BANCARIA <> "" Then
            txtIBAN = .getCUENTA_BANCARIA
        End If
        'M1334-F
        cmbFP.BoundText = .getFP_ID
        cmbVencimiento.BoundText = .getVENCIMIENTO_ID
        cmbBanco.BoundText = .getBANCO_ID
        ' Pais
        Dim opais As New clsPais
        opais.CargarPais (.getPAIS_ID)
        cmbPais.BoundText = opais.getNOMBRE
        cmbPais.Text = opais.getNOMBRE
        Set opais = Nothing
        ' Provincia
        Dim oProvincia As New clsProvincias
        oProvincia.CargarProvincia (.getPROVINCIA_ID)
        cmbProvincia.BoundText = oProvincia.getNOMBRE
        cmbProvincia.Text = oProvincia.getNOMBRE
        Set oProvincia = Nothing
        ' Municipio
        Dim oMunicipio As New clsMunicipios
        oMunicipio.CargarMunicipio (.getMUNICIPIO_ID)
        cmbMunicipios.BoundText = oMunicipio.getNOMBRE
        cmbMunicipios.Text = oMunicipio.getNOMBRE
        Set oMunicipio = Nothing
        If .getANULADO = 0 Then
            chkAnulado.Value = Unchecked
        Else
            chkAnulado.Value = Checked
        End If
    End With
    ' Adjuntos
    Dim oPA As New clsProveedores_Adjuntos
    Dim rs As ADODB.Recordset
    Set rs = oPA.Listado(PK)
    If rs.RecordCount > 0 Then
        Do
            With lista.ListItems.Add(, , rs("ORDEN"))
                .SubItems(1) = rs("DESCRIPCION")
                .SubItems(2) = rs("RUTA")
            End With
            rs.MoveNext
        Loop Until rs.EOF
    End If
    ' Servicios
    cargar_servicios
    cargarPto
    Set oPA = Nothing
    Set oProveedor = Nothing
    Exit Sub
fallo:
    MsgBox "Error al consultar los datos del proveedor.", vbCritical, Err.Description
End Sub

Public Sub desbloquear_controles()
    Dim i As Integer
    For i = 1 To 13
        txtdatos(i).Locked = False
    Next
    cmbMunicipios.Locked = False
    cmbProvincia.Locked = False
    cmbPais.Locked = False
End Sub

Public Function mover_datos() As clsProveedor
    On Error GoTo fallo
    Dim oProveedor As New clsProveedor
    With oProveedor
        'M1339-I
        .setEMAIL_FACTURACION = txtdatos(0)
        'M1339-F
        .setNOMBRE = txtdatos(1)
        .setDIRECCION = txtdatos(2)
        If txtdatos(3) <> "" Then
            .setCOD_POSTAL = CLng(txtdatos(3))
        Else
            .setCOD_POSTAL = 0
        End If
        .setCIF = txtdatos(6)
        .setTELEFONO = txtdatos(4)
        .setFAX = txtdatos(5)
        .setRESPONSABLE = txtdatos(8)
        .setTIPO = "" ' Ojo
        .setTIPO = "0" ' JONATHAN.2010.05.13
        .setEMAIL = txtdatos(7)
        .setOBSERVACIONES = txtdatos(13)
        'TESORERIA-I
        .setCC = txtdatos(14)
        .setCC_RETENCION = txtdatos(9)
        'TESORERIA-F
        .setWEB = txtdatos(12)
        .setFP_ID = cmbFP.BoundText
        .setDIAS_ENTREGA = txtdatos(11)
        If cmbVencimiento.Text = "" Then
            .setVENCIMIENTO_ID = 0
        Else
            .setVENCIMIENTO_ID = cmbVencimiento.BoundText
        End If
        If cmbBanco.Text = "" Then
            .setBANCO_ID = 0
        Else
            .setBANCO_ID = cmbBanco.BoundText
        End If
        'E0200-I
        If chkSubcontrata.Value = Unchecked Then
            .setES_SUBCONTRATA = 0
        Else
            .setES_SUBCONTRATA = 1
        End If
        'E0200-F
        'MXXXX-I
        If chkFormador.Value = Unchecked Then
            .setES_FORMADOR = 0
        Else
            .setES_FORMADOR = 1
        End If
        .setINTRA = chkIntra.Value
        .setEXTRA = chkExtra.Value
        'MXXXX-F
        ' Pais
        If cmbPais.Text <> "" Then
            If IsNumeric(cmbPais.BoundText) Then
                .setPAIS_ID = cmbPais.BoundText
            Else
                Dim opais As New clsPais
                Dim pais As Long
                pais = opais.buscar(cmbPais.Text)
                If pais = 0 Then
                    opais.setNOMBRE = cmbPais.Text
                    .setPAIS_ID = opais.Insertar
                Else
                    .setPAIS_ID = pais
                End If
            End If
        End If
        ' Provincia
        If cmbProvincia.Text <> "" Then
            If IsNumeric(cmbProvincia.BoundText) Then
                .setPROVINCIA_ID = cmbProvincia.BoundText
            Else
                Dim oprov As New clsProvincias
                Dim PROVINCIA As Long
                PROVINCIA = oprov.buscar(cmbProvincia.Text)
                If PROVINCIA = 0 Then
                    oprov.setPAIS_ID = .getPAIS_ID
                    oprov.setNOMBRE = cmbProvincia.Text
                    .setPROVINCIA_ID = oprov.Insertar
                Else
                    .setPROVINCIA_ID = PROVINCIA
                End If
            End If
        End If
        ' Municipio
        If cmbMunicipios.Text <> "" Then
            If IsNumeric(cmbMunicipios.BoundText) Then
                .setMUNICIPIO_ID = cmbMunicipios.BoundText
            Else
                Dim omun As New clsMunicipios
                Dim municipio As Long
                municipio = omun.buscar(cmbMunicipios.Text)
                If municipio = 0 Then
                    omun.setPROVINCIA_ID = .getPROVINCIA_ID
                    omun.setNOMBRE = cmbMunicipios.Text
                    .setMUNICIPIO_ID = omun.Insertar
                Else
                    .setMUNICIPIO_ID = municipio
                End If
            End If
        End If
        'M1334-I
        'CUENTA BANCARIA
'        If chkIBAN.value = Unchecked Then
'            .setES_CUENTA_ESPANYA = 0
'        Else
'            .setES_CUENTA_ESPANYA = 1
'        End If
'        .setCUENTA_BANCARIA = Trim(txtCuentaIBAN.Text)
        .setBANCO = txtdatos(10)
        .setCUENTA_BANCARIA = txtIBAN.Text
        'M1334-F
        .setANULADO = chkAnulado.Value
    End With
    Set mover_datos = oProveedor
    Set oProveedor = Nothing
    Exit Function
fallo:
    MsgBox "Error al mover los datos del proveedor.", vbCritical, Err.Description
End Function
Public Sub cargar_paises()
    Dim opais As New clsPais
    Set cmbPais.RowSource = opais.Listado  'recorset devuelto por la funcion
    cmbPais.ListField = "nombre" 'campo que veo
    cmbPais.DataField = "nombre" 'campo asociado
    cmbPais.BoundColumn = "id_pais" 'lo que realmente envia
    Set opais = Nothing
End Sub
Public Sub cargar_provincias()
'    cmbProvincia.Text = ""
    If cmbPais.Text <> "" Then
     If IsNumeric(cmbPais.BoundText) Then
        Dim oProvincia As New clsProvincias
        Set cmbProvincia.RowSource = oProvincia.Listado(CInt(cmbPais.BoundText))  'recorset devuelto por la funcion
        cmbProvincia.ListField = "nombre" 'campo que veo
        cmbProvincia.DataField = "nombre" 'campo asociado
        cmbProvincia.BoundColumn = "id_provincia" 'lo que realmente envia
        Set oProvincia = Nothing
     End If
    End If
End Sub
Public Sub cargar_municipios()
'    cmbMunicipios.Text = ""
    If cmbProvincia.Text <> "" Then
     If IsNumeric(cmbProvincia.BoundText) Then
        Dim omuni As New clsMunicipios
        Set cmbMunicipios.RowSource = omuni.Listado(CInt(cmbProvincia.BoundText))
        cmbMunicipios.ListField = "nombre" 'campo que veo
        cmbMunicipios.DataField = "nombre" 'campo asociado
        cmbMunicipios.BoundColumn = "id_municipio" 'lo que realmente envia
        Set omuni = Nothing
     End If
    End If
End Sub


Private Sub cabecera()
    With lista.ColumnHeaders
         .Add , , "ORDEN", 1, lvwColumnLeft
         .Add , , "Descripción", lista.Width, lvwColumnLeft
         .Add , , "Ruta", 1, lvwColumnLeft
    End With
End Sub

Private Sub insertar_adjuntos()
    ' Evidencias
    Dim oPA As New clsProveedores_Adjuntos
    oPA.Eliminar PK
    Dim i As Integer
    For i = 1 To lista.ListItems.Count
        With oPA
            .setPROVEEDOR_ID = PK
            .setDESCRIPCION = lista.ListItems(i).SubItems(1)
            .setRUTA = lista.ListItems(i).SubItems(2)
            .setORDEN = i
            .Insertar
        End With
    Next
    Set oPA = Nothing
End Sub
Private Sub insertar_servicios()
    ' Evidencias
    Dim oPS As New clsProveedores_servicios
   On Error GoTo insertar_servicios_Error

    oPS.Eliminar PK
    Dim i As Integer
    For i = X.LowerBound(1) To X.UpperBound(1)
        If Trim(X.Value(i, COLS.servicio)) <> "" Then
            With oPS
                .setPROVEEDOR_ID = PK
                .setORDEN = i
                .setSERVICIO_ID = X.Value(i, COLS.ID_SERVICIO)
                .setTIPO = X.Value(i, COLS.tipo)
                .setMETODO_ID = X.Value(i, COLS.ID_METODO)
                .setFECHA_INICIAL = Format(X.Value(i, COLS.finicial), "yyyy-mm-dd")
                If IsEmpty(X.Value(i, COLS.fAprobacion)) Or (X.Value(i, COLS.fAprobacion) = "") Then
                    .setFECHA_APROBACION = "9999-12-31"
                Else
                    .setFECHA_APROBACION = Format(X.Value(i, COLS.fAprobacion), "yyyy-mm-dd")
                End If
                .setOBSERVACIONES = X.Value(i, COLS.OBSERVACIONES)
                .Insertar
            End With
        End If
    Next
    Set oPS = Nothing

   On Error GoTo 0
   Exit Sub

insertar_servicios_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure insertar_servicios of Formulario frmProveedores_Detalle"
End Sub
Private Sub insertarPto()
    Dim oPP As New clsProveedores_pto
   On Error GoTo insertar_servicios_Error

    oPP.Eliminar PK
    Dim i As Integer
    For i = xPto.LowerBound(1) To xPto.UpperBound(1)
        If Trim(xPto.Value(i, COLSPTO.ANO)) <> "" Then
            With oPP
                .setPROVEEDOR_ID = PK
                .setANO = xPto.Value(i, COLSPTO.ANO)
                If xPto.Value(i, COLSPTO.PRESUPUESTO) <> "" Then
                    .setPRESUPUESTO = moneda_bd(xPto.Value(i, COLSPTO.PRESUPUESTO))
                Else
                    .setPRESUPUESTO = moneda_bd("0")
                End If
                If xPto.Value(i, COLSPTO.REAL) <> "" Then
                    .setREAL = moneda_bd(xPto.Value(i, COLSPTO.REAL))
                Else
                    .setREAL = moneda_bd("0")
                End If
                .Insertar
            End With
        End If
    Next
    Set oPP = Nothing

   On Error GoTo 0
   Exit Sub

insertar_servicios_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure insertar_servicios of Formulario frmProveedores_Detalle"
End Sub

Private Sub cargar_combos()
'    cargar_combo cmbFP, New clsFP
    cargar_combo cmbBanco, New clsBancos
    Dim rs As ADODB.Recordset
    ' Servicios
    Dim oDeco As New clsDecodificadora
    oDeco.cargar_combo cmbFP, DECODIFICADORA.DECODIFICADORA_PROVEEDORES_FP
    oDeco.cargar_combo cmbVencimiento, DECODIFICADORA.DECODIFICADORA_PROVEEDORES_VENCIMIENTOS
    
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
    X.ReDim 0, filas, 0, Col
    X.Clear
    Set grid.Array = X
    grid.Refresh
    ' Grid Pto
    xPto.ReDim 0, filasPto, 0, ColPto
    xPto.Clear
    Set gridPto.Array = xPto
    gridPto.Refresh
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
            X(i, COLS.servicio) = CStr(oDeco.getDESCRIPCION)
            X(i, COLS.tipo) = CStr(rs("TIPO"))
            oDeco.Carga_valor DECODIFICADORA.PROVEEDORES_METODOS, rs("METODO_ID")
            X(i, COLS.METODO) = CStr(oDeco.getDESCRIPCION)
            X(i, COLS.finicial) = CStr(Format(rs("FECHA_INICIAL"), "dd-mm-yyyy"))
            If Format(rs("FECHA_APROBACION"), "yyyy-mm-dd") <> "9999-12-31" Then
                X(i, COLS.fAprobacion) = CStr(Format(rs("FECHA_APROBACION"), "dd-mm-yyyy"))
            End If
            X(i, COLS.OBSERVACIONES) = CStr(rs("OBSERVACIONES"))
            X(i, COLS.ID_SERVICIO) = CStr(rs("SERVICIO_ID"))
            X(i, COLS.ID_METODO) = CStr(rs("METODO_ID"))
            i = i + 1
            rs.MoveNext
        Loop Until rs.EOF
    End If
'    numero_servicios
    Set rs = Nothing
    Set oDeco = Nothing
End Sub
Private Sub cargarPto()
    Dim oPP As New clsProveedores_pto
    Dim oPF As New clsProveedores_Facturas
    Dim rs As ADODB.Recordset
    Dim ANNO As Integer
    Dim i As Integer
    Dim fila As Integer
    fila = 0
    ANNO = Year(Date)
    For i = ANNO To 2015 Step -1
        xPto(fila, COLSPTO.ANO) = CStr(i)
        ' Recuperar Presupuesto
        Set rs = oPP.Listado(PK, i)
        If rs.RecordCount > 0 Then
            xPto(fila, COLSPTO.PRESUPUESTO) = CStr(rs("PRESUPUESTO"))
        Else
            xPto(fila, COLSPTO.PRESUPUESTO) = CStr("0")
        End If
        ' Real
        xPto(fila, COLSPTO.REAL) = oPF.importeProveedor(PK, i)
        fila = fila + 1
    Next
    Set rs = Nothing
End Sub

Private Sub numero_servicios()
    Dim cont As Integer
    cont = 0
    Dim i As Integer
    For i = 0 To filas
        If Trim(X.Value(i, COLS.servicio)) <> "" Then
            cont = cont + 1
        End If
    Next
    lblServicios.Caption = "Número de Servicios : " & cont
End Sub
