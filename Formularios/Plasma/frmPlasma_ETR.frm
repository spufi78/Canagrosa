VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{CEA3B7F6-2847-4E5E-A551-DB7A62489D44}#46.0#0"; "miCombo.ocx"
Object = "{77DDF307-D82B-4757-8B3A-106EC9D75D4B}#8.0#0"; "tdbg8.ocx"
Begin VB.Form frmPlasma_ETR 
   BackColor       =   &H00C0C0C0&
   Caption         =   "Registro de Resultados Muestra de Plasma"
   ClientHeight    =   10440
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   13725
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmPlasma_ETR.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   10440
   ScaleWidth      =   13725
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtDurezaEspesor 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   9090
      MaxLength       =   255
      TabIndex        =   53
      Top             =   10035
      Visible         =   0   'False
      Width           =   2085
   End
   Begin VB.TextBox txtUnidades 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   2520
      TabIndex        =   50
      Top             =   9810
      Visible         =   0   'False
      Width           =   1815
   End
   Begin Geslab.ControlPanelXP cpReactivos 
      Height          =   3435
      Left            =   6885
      TabIndex        =   21
      Top             =   2745
      Width           =   6810
      _ExtentX        =   12012
      _ExtentY        =   6059
      Caption         =   "Reactivos Utilizados en la Muestra"
      BackColor       =   12632256
      TextColor       =   0
      HeaderColor     =   8421504
      Object.Height          =   3435
      Begin VB.Frame frmReactivos 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   0  'None
         Caption         =   "Reactivos"
         ForeColor       =   &H80000008&
         Height          =   2940
         Left            =   45
         TabIndex        =   22
         Top             =   405
         Width           =   6630
         Begin VB.CommandButton cmdAnadirReactivo 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Añadir"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   750
            Left            =   5670
            Style           =   1  'Graphical
            TabIndex        =   24
            Tag             =   "Añade campo o modifica el campo existente con el mismo nombre"
            Top             =   1035
            Width           =   915
         End
         Begin VB.CommandButton cmdEliminarReactivo 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Eliminar"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   795
            Left            =   5670
            Style           =   1  'Graphical
            TabIndex        =   23
            Tag             =   "Elimina el campo seleccionado"
            Top             =   180
            Width           =   915
         End
         Begin MSComctlLib.ListView listaReactivos 
            Height          =   1830
            Left            =   45
            TabIndex        =   25
            Top             =   135
            Width           =   5445
            _ExtentX        =   9604
            _ExtentY        =   3228
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
         Begin pryCombo.miCombo cmbReactivos 
            Height          =   330
            Left            =   765
            TabIndex        =   26
            Top             =   2115
            Width           =   5280
            _ExtentX        =   9313
            _ExtentY        =   582
         End
         Begin pryCombo.miCombo cmbReactivosInternos 
            Height          =   330
            Left            =   765
            TabIndex        =   27
            Top             =   2475
            Width           =   5280
            _ExtentX        =   9313
            _ExtentY        =   582
         End
         Begin VB.Label lblCampos 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C0C0&
            Caption         =   "Externo"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   5
            Left            =   90
            TabIndex        =   29
            Top             =   2160
            Width           =   540
         End
         Begin VB.Label lblCampos 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C0C0&
            Caption         =   "Interno"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   8
            Left            =   90
            TabIndex        =   28
            Top             =   2520
            Width           =   495
         End
      End
   End
   Begin Geslab.ControlPanelXP ControlPanelXP2 
      Height          =   3435
      Left            =   45
      TabIndex        =   14
      Top             =   2745
      Width           =   6810
      _ExtentX        =   12012
      _ExtentY        =   6059
      Caption         =   "Equipos Utilizados en la Muestra"
      BackColor       =   12632256
      TextColor       =   0
      HeaderColor     =   8421504
      Object.Height          =   3435
      Begin VB.Frame frmEquipos 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   2985
         Left            =   90
         TabIndex        =   15
         Top             =   405
         Width           =   6585
         Begin VB.CommandButton cmdVerificacion 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Verificación"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   765
            Left            =   5670
            Style           =   1  'Graphical
            TabIndex        =   31
            Tag             =   "Añade campo o modifica el campo existente con el mismo nombre"
            Top             =   1920
            Width           =   915
         End
         Begin VB.CommandButton cmdAnadirEquipo 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Añadir"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   765
            Left            =   5670
            Style           =   1  'Graphical
            TabIndex        =   17
            Tag             =   "Añade campo o modifica el campo existente con el mismo nombre"
            Top             =   1110
            Width           =   915
         End
         Begin VB.CommandButton cmdEliminarEquipo 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Eliminar"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   810
            Left            =   5670
            Style           =   1  'Graphical
            TabIndex        =   16
            Tag             =   "Elimina el campo seleccionado"
            Top             =   270
            Width           =   915
         End
         Begin MSComctlLib.ListView listaEquipos 
            Height          =   2325
            Left            =   0
            TabIndex        =   18
            Top             =   270
            Width           =   5580
            _ExtentX        =   9843
            _ExtentY        =   4101
            View            =   3
            LabelEdit       =   1
            LabelWrap       =   -1  'True
            HideSelection   =   0   'False
            AllowReorder    =   -1  'True
            Checkboxes      =   -1  'True
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
         Begin pryCombo.miCombo cmbEquipos 
            Height          =   330
            Left            =   0
            TabIndex        =   19
            Top             =   2610
            Width           =   5595
            _ExtentX        =   9869
            _ExtentY        =   582
         End
         Begin VB.Label Label2 
            BackColor       =   &H00C0C0C0&
            Caption         =   "Marque los equipos que deben salir en el informe"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   195
            Index           =   0
            Left            =   0
            TabIndex        =   20
            Top             =   45
            Width           =   4335
         End
      End
   End
   Begin VB.Frame frmRockwell 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      ForeColor       =   &H80000008&
      Height          =   3255
      Left            =   45
      TabIndex        =   45
      Top             =   6210
      Width           =   13650
      Begin VB.TextBox txtSD 
         Alignment       =   2  'Center
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
         ForeColor       =   &H000000FF&
         Height          =   345
         Left            =   11565
         Locked          =   -1  'True
         MaxLength       =   255
         TabIndex        =   48
         Top             =   2835
         Width           =   1950
      End
      Begin VB.TextBox txtDurezaAverage 
         Alignment       =   2  'Center
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
         ForeColor       =   &H000000FF&
         Height          =   345
         Left            =   8595
         Locked          =   -1  'True
         MaxLength       =   255
         TabIndex        =   46
         Top             =   2835
         Width           =   1680
      End
      Begin TrueDBGrid80.TDBGrid grid 
         Height          =   2595
         Left            =   45
         TabIndex        =   55
         Top             =   180
         Width           =   13560
         _ExtentX        =   23918
         _ExtentY        =   4577
         _LayoutType     =   4
         _RowHeight      =   16
         _WasPersistedAsPixels=   0
         Columns(0)._VlistStyle=   0
         Columns(0)._MaxComboItems=   5
         Columns(0).Caption=   "Identification"
         Columns(0).DataField=   ""
         Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(1)._VlistStyle=   0
         Columns(1)._MaxComboItems=   5
         Columns(1).Caption=   "Lado 1 (in)"
         Columns(1).DataField=   ""
         Columns(1).NumberFormat=   "0.000"
         Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(2)._VlistStyle=   0
         Columns(2)._MaxComboItems=   5
         Columns(2).Caption=   "Lado 2 (in)"
         Columns(2).DataField=   ""
         Columns(2).NumberFormat=   "0.000"
         Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(3)._VlistStyle=   0
         Columns(3)._MaxComboItems=   5
         Columns(3).Caption=   "Valor Obtenido (lb)"
         Columns(3).DataField=   ""
         Columns(3).NumberFormat=   "General Number"
         Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(4)._VlistStyle=   0
         Columns(4)._MaxComboItems=   5
         Columns(4).Caption=   "Area (in)"
         Columns(4).DataField=   ""
         Columns(4).NumberFormat=   "0.000"
         Columns(4)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(5)._VlistStyle=   0
         Columns(5)._MaxComboItems=   5
         Columns(5).Caption=   "Resultado (psi)"
         Columns(5).DataField=   ""
         Columns(5)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(6)._VlistStyle=   0
         Columns(6)._MaxComboItems=   5
         Columns(6).Caption=   "Resultado (MPA)"
         Columns(6).DataField=   ""
         Columns(6).NumberFormat=   "0.0"
         Columns(6)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns.Count   =   7
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
         Splits(0)._ColumnProps(0)=   "Columns.Count=7"
         Splits(0)._ColumnProps(1)=   "Column(0).Width=5821"
         Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
         Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=5741"
         Splits(0)._ColumnProps(4)=   "Column(0)._EditAlways=0"
         Splits(0)._ColumnProps(5)=   "Column(0)._ColStyle=1"
         Splits(0)._ColumnProps(6)=   "Column(0).Order=1"
         Splits(0)._ColumnProps(7)=   "Column(1).Width=2963"
         Splits(0)._ColumnProps(8)=   "Column(1).DividerColor=0"
         Splits(0)._ColumnProps(9)=   "Column(1)._WidthInPix=2884"
         Splits(0)._ColumnProps(10)=   "Column(1)._EditAlways=0"
         Splits(0)._ColumnProps(11)=   "Column(1)._ColStyle=1"
         Splits(0)._ColumnProps(12)=   "Column(1).Order=2"
         Splits(0)._ColumnProps(13)=   "Column(2).Width=2937"
         Splits(0)._ColumnProps(14)=   "Column(2).DividerColor=0"
         Splits(0)._ColumnProps(15)=   "Column(2)._WidthInPix=2858"
         Splits(0)._ColumnProps(16)=   "Column(2)._EditAlways=0"
         Splits(0)._ColumnProps(17)=   "Column(2)._ColStyle=1"
         Splits(0)._ColumnProps(18)=   "Column(2).Order=3"
         Splits(0)._ColumnProps(19)=   "Column(3).Width=2990"
         Splits(0)._ColumnProps(20)=   "Column(3).DividerColor=0"
         Splits(0)._ColumnProps(21)=   "Column(3)._WidthInPix=2910"
         Splits(0)._ColumnProps(22)=   "Column(3)._EditAlways=0"
         Splits(0)._ColumnProps(23)=   "Column(3)._ColStyle=1"
         Splits(0)._ColumnProps(24)=   "Column(3).Order=4"
         Splits(0)._ColumnProps(25)=   "Column(4).Width=2619"
         Splits(0)._ColumnProps(26)=   "Column(4).DividerColor=0"
         Splits(0)._ColumnProps(27)=   "Column(4)._WidthInPix=2540"
         Splits(0)._ColumnProps(28)=   "Column(4)._EditAlways=0"
         Splits(0)._ColumnProps(29)=   "Column(4)._ColStyle=8193"
         Splits(0)._ColumnProps(30)=   "Column(4).AllowFocus=0"
         Splits(0)._ColumnProps(31)=   "Column(4).Order=5"
         Splits(0)._ColumnProps(32)=   "Column(5).Width=2805"
         Splits(0)._ColumnProps(33)=   "Column(5).DividerColor=0"
         Splits(0)._ColumnProps(34)=   "Column(5)._WidthInPix=2725"
         Splits(0)._ColumnProps(35)=   "Column(5)._EditAlways=0"
         Splits(0)._ColumnProps(36)=   "Column(5)._ColStyle=8193"
         Splits(0)._ColumnProps(37)=   "Column(5).AllowFocus=0"
         Splits(0)._ColumnProps(38)=   "Column(5).Order=6"
         Splits(0)._ColumnProps(39)=   "Column(6).Width=2725"
         Splits(0)._ColumnProps(40)=   "Column(6).DividerColor=0"
         Splits(0)._ColumnProps(41)=   "Column(6)._WidthInPix=2646"
         Splits(0)._ColumnProps(42)=   "Column(6)._EditAlways=0"
         Splits(0)._ColumnProps(43)=   "Column(6)._ColStyle=8193"
         Splits(0)._ColumnProps(44)=   "Column(6).AllowFocus=0"
         Splits(0)._ColumnProps(45)=   "Column(6).Order=7"
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
         TabAction       =   2
         WrapCellPointer =   -1  'True
         MultipleLines   =   0
         EmptyRows       =   -1  'True
         CellTipsWidth   =   0
         MultiSelect     =   2
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
         _StyleDefs(36)  =   "Splits(0).Columns(0).Style:id=24,.parent=11,.alignment=2,.bgcolor=&HDEEDFA&"
         _StyleDefs(37)  =   ":id=24,.locked=0,.bold=-1,.fontsize=825,.italic=0,.underline=0,.strikethrough=0"
         _StyleDefs(38)  =   ":id=24,.charset=0"
         _StyleDefs(39)  =   ":id=24,.fontname=MS Sans Serif"
         _StyleDefs(40)  =   "Splits(0).Columns(0).HeadingStyle:id=21,.parent=12"
         _StyleDefs(41)  =   "Splits(0).Columns(0).FooterStyle:id=22,.parent=13"
         _StyleDefs(42)  =   "Splits(0).Columns(0).EditorStyle:id=23,.parent=15"
         _StyleDefs(43)  =   "Splits(0).Columns(1).Style:id=28,.parent=11,.alignment=2,.bgcolor=&HDEEDFA&"
         _StyleDefs(44)  =   ":id=28,.locked=0"
         _StyleDefs(45)  =   "Splits(0).Columns(1).HeadingStyle:id=25,.parent=12"
         _StyleDefs(46)  =   "Splits(0).Columns(1).FooterStyle:id=26,.parent=13"
         _StyleDefs(47)  =   "Splits(0).Columns(1).EditorStyle:id=27,.parent=15"
         _StyleDefs(48)  =   "Splits(0).Columns(2).Style:id=62,.parent=11,.alignment=2"
         _StyleDefs(49)  =   "Splits(0).Columns(2).HeadingStyle:id=59,.parent=12"
         _StyleDefs(50)  =   "Splits(0).Columns(2).FooterStyle:id=60,.parent=13"
         _StyleDefs(51)  =   "Splits(0).Columns(2).EditorStyle:id=61,.parent=15"
         _StyleDefs(52)  =   "Splits(0).Columns(3).Style:id=54,.parent=11,.alignment=2,.bgcolor=&HDEEDFA&"
         _StyleDefs(53)  =   ":id=54,.locked=0"
         _StyleDefs(54)  =   "Splits(0).Columns(3).HeadingStyle:id=51,.parent=12"
         _StyleDefs(55)  =   "Splits(0).Columns(3).FooterStyle:id=52,.parent=13"
         _StyleDefs(56)  =   "Splits(0).Columns(3).EditorStyle:id=53,.parent=15"
         _StyleDefs(57)  =   "Splits(0).Columns(4).Style:id=32,.parent=11,.alignment=2,.bgcolor=&HD7D7D7&"
         _StyleDefs(58)  =   ":id=32,.locked=-1"
         _StyleDefs(59)  =   "Splits(0).Columns(4).HeadingStyle:id=29,.parent=12"
         _StyleDefs(60)  =   "Splits(0).Columns(4).FooterStyle:id=30,.parent=13"
         _StyleDefs(61)  =   "Splits(0).Columns(4).EditorStyle:id=31,.parent=15"
         _StyleDefs(62)  =   "Splits(0).Columns(5).Style:id=36,.parent=11,.alignment=2,.bgcolor=&HD7D7D7&"
         _StyleDefs(63)  =   ":id=36,.locked=-1"
         _StyleDefs(64)  =   "Splits(0).Columns(5).HeadingStyle:id=33,.parent=12"
         _StyleDefs(65)  =   "Splits(0).Columns(5).FooterStyle:id=34,.parent=13"
         _StyleDefs(66)  =   "Splits(0).Columns(5).EditorStyle:id=35,.parent=15"
         _StyleDefs(67)  =   "Splits(0).Columns(6).Style:id=58,.parent=11,.alignment=2,.bgcolor=&HD7D7D7&"
         _StyleDefs(68)  =   ":id=58,.locked=-1"
         _StyleDefs(69)  =   "Splits(0).Columns(6).HeadingStyle:id=55,.parent=12"
         _StyleDefs(70)  =   "Splits(0).Columns(6).FooterStyle:id=56,.parent=13"
         _StyleDefs(71)  =   "Splits(0).Columns(6).EditorStyle:id=57,.parent=15"
         _StyleDefs(72)  =   "Named:id=37:Normal"
         _StyleDefs(73)  =   ":id=37,.parent=0,.alignment=3,.bgcolor=&H80000018&,.bgpicMode=2,.bold=0"
         _StyleDefs(74)  =   ":id=37,.fontsize=825,.italic=0,.underline=0,.strikethrough=0,.charset=0"
         _StyleDefs(75)  =   ":id=37,.fontname=MS Sans Serif"
         _StyleDefs(76)  =   "Named:id=38:Heading"
         _StyleDefs(77)  =   ":id=38,.parent=37,.valignment=2,.bgcolor=&H80000016&,.fgcolor=&H80000012&"
         _StyleDefs(78)  =   ":id=38,.wraptext=-1,.bold=0,.fontsize=825,.italic=0,.underline=0"
         _StyleDefs(79)  =   ":id=38,.strikethrough=0,.charset=0"
         _StyleDefs(80)  =   ":id=38,.fontname=MS Sans Serif"
         _StyleDefs(81)  =   "Named:id=39:Footing"
         _StyleDefs(82)  =   ":id=39,.parent=37,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
         _StyleDefs(83)  =   "Named:id=40:Selected"
         _StyleDefs(84)  =   ":id=40,.parent=37,.bgcolor=&H8080FF&,.fgcolor=&H0&,.bold=0,.fontsize=825"
         _StyleDefs(85)  =   ":id=40,.italic=0,.underline=0,.strikethrough=0,.charset=0"
         _StyleDefs(86)  =   ":id=40,.fontname=MS Sans Serif"
         _StyleDefs(87)  =   "Named:id=41:Caption"
         _StyleDefs(88)  =   ":id=41,.parent=38,.alignment=2"
         _StyleDefs(89)  =   "Named:id=42:HighlightRow"
         _StyleDefs(90)  =   ":id=42,.parent=37,.bgcolor=&H80000008&,.fgcolor=&H80000005&"
         _StyleDefs(91)  =   "Named:id=43:EvenRow"
         _StyleDefs(92)  =   ":id=43,.parent=37,.bgcolor=&HFFFFFF&,.wraptext=-1"
         _StyleDefs(93)  =   "Named:id=44:OddRow"
         _StyleDefs(94)  =   ":id=44,.parent=37,.bgcolor=&HD5ECF9&"
         _StyleDefs(95)  =   "Named:id=47:RecordSelector"
         _StyleDefs(96)  =   ":id=47,.parent=38"
         _StyleDefs(97)  =   "Named:id=50:FilterBar"
         _StyleDefs(98)  =   ":id=50,.parent=37"
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Media (Mpa)"
         Height          =   195
         Index           =   20
         Left            =   10440
         TabIndex        =   49
         Top             =   2880
         Width           =   1065
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Media (psi)"
         Height          =   195
         Index           =   6
         Left            =   7560
         TabIndex        =   47
         Top             =   2880
         Width           =   945
      End
   End
   Begin VB.Frame Frame5 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   "RESULT"
      ForeColor       =   &H80000008&
      Height          =   735
      Left            =   4545
      TabIndex        =   42
      Top             =   9540
      Width           =   4425
      Begin VB.CheckBox chkResult 
         BackColor       =   &H00C0C0C0&
         Caption         =   "chkResult"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   225
         TabIndex        =   43
         Top             =   270
         Value           =   1  'Checked
         Width           =   240
      End
      Begin VB.Label lblResult 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C0C0&
         Caption         =   "PASS"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00008000&
         Height          =   375
         Left            =   540
         TabIndex        =   44
         Top             =   225
         Width           =   3390
      End
   End
   Begin VB.Frame Frame6 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   "SPECIMEN ID AND DESCRIPTION"
      ForeColor       =   &H80000008&
      Height          =   2370
      Left            =   45
      TabIndex        =   33
      Top             =   360
      Width           =   13650
      Begin VB.TextBox txtDatos 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   50
         Left            =   1350
         MaxLength       =   255
         TabIndex        =   2
         Top             =   990
         Width           =   11085
      End
      Begin VB.TextBox txtDatos 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   51
         Left            =   1350
         MaxLength       =   255
         TabIndex        =   3
         Top             =   1665
         Width           =   2715
      End
      Begin VB.TextBox txtDatos 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   54
         Left            =   1350
         MaxLength       =   255
         TabIndex        =   6
         Top             =   1980
         Width           =   2715
      End
      Begin VB.TextBox txtDatos 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   52
         Left            =   5580
         MaxLength       =   255
         TabIndex        =   4
         Top             =   1665
         Width           =   2760
      End
      Begin VB.TextBox txtDatos 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   55
         Left            =   5580
         MaxLength       =   255
         TabIndex        =   7
         Top             =   1980
         Width           =   2760
      End
      Begin VB.TextBox txtDatos 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   53
         Left            =   9630
         MaxLength       =   255
         TabIndex        =   5
         Top             =   1665
         Width           =   2805
      End
      Begin pryCombo.miCombo cmbProcess 
         Height          =   345
         Left            =   1350
         TabIndex        =   0
         Top             =   270
         Width           =   12150
         _ExtentX        =   21431
         _ExtentY        =   609
      End
      Begin pryCombo.miCombo cmbCustomer 
         Height          =   345
         Left            =   1350
         TabIndex        =   1
         Top             =   630
         Width           =   12150
         _ExtentX        =   21431
         _ExtentY        =   609
      End
      Begin pryCombo.miCombo cmbnatype 
         Height          =   345
         Left            =   1350
         TabIndex        =   51
         Top             =   1305
         Width           =   12150
         _ExtentX        =   21431
         _ExtentY        =   609
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Nº AND TYPE"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   54
         Left            =   135
         TabIndex        =   52
         Top             =   1350
         Width           =   1035
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "S/N:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   3
         Left            =   135
         TabIndex        =   41
         Top             =   2025
         Width           =   345
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "SPECIMEN ID"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   2
         Left            =   135
         TabIndex        =   40
         Top             =   1035
         Width           =   1035
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "P/N:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   1
         Left            =   135
         TabIndex        =   39
         Top             =   1725
         Width           =   345
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "PROCESS"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   11
         Left            =   135
         TabIndex        =   38
         Top             =   345
         Width           =   765
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "CUSTOMER"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   14
         Left            =   135
         TabIndex        =   37
         Top             =   705
         Width           =   915
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "PRODUCT S/N:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   15
         Left            =   4185
         TabIndex        =   36
         Top             =   2025
         Width           =   1185
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "PRODUCT TYPE:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   16
         Left            =   4185
         TabIndex        =   35
         Top             =   1725
         Width           =   1305
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "MODULE S/N:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   17
         Left            =   8505
         TabIndex        =   34
         Top             =   1710
         Width           =   1080
      End
   End
   Begin VB.CommandButton cmdImagen 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Imagenes"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   840
      Left            =   90
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   9540
      Width           =   1140
   End
   Begin VB.CommandButton cmdObservador 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Observador"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   840
      Left            =   1230
      Style           =   1  'Graphical
      TabIndex        =   9
      Tag             =   "Añade campo o modifica el campo existente con el mismo nombre"
      Top             =   9540
      Visible         =   0   'False
      Width           =   1140
   End
   Begin VB.CheckBox chkModificar 
      Caption         =   "Permiso Modificar Cerrada"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   9315
      TabIndex        =   32
      Top             =   9495
      Visible         =   0   'False
      Width           =   1230
   End
   Begin VB.CheckBox chkDuplicada 
      Caption         =   "Duplicada"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   9315
      TabIndex        =   30
      Top             =   9720
      Visible         =   0   'False
      Width           =   1050
   End
   Begin VB.CommandButton cmdSalir 
      BackColor       =   &H00E0E0E0&
      Caption         =   "ESC-Salir"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   840
      Left            =   12510
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   9540
      Width           =   1140
   End
   Begin VB.CommandButton cmdok 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Aceptar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   840
      Left            =   11340
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   9540
      Width           =   1140
   End
   Begin VB.Label lblCampos 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      Caption         =   "THICKNESS"
      Height          =   195
      Index           =   19
      Left            =   5220
      TabIndex        =   54
      Top             =   10170
      Visible         =   0   'False
      Width           =   1065
   End
   Begin VB.Label lblCerrada 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   345
      Left            =   11925
      TabIndex        =   13
      Top             =   0
      Width           =   1770
   End
   Begin VB.Label lbltitulo 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Caption         =   "Resultados de Muestra de Plasma"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   330
      Left            =   45
      TabIndex        =   12
      Top             =   0
      Width           =   13725
   End
End
Attribute VB_Name = "frmPlasma_ETR"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public PK As Long

Dim xTRACCION As New XArrayDB
Const filasGrid As Integer = 25
Const ColGrid As Integer = 7
Private Enum ColsGrid
    IDENTIFICATION = 0
    LADO1 = 1
    LADO2 = 2
    VALOR_OBTENIDO = 3
    area = 4
    RESULTADO_PSI = 5
    RESULTADO_MPA = 6
End Enum
Private Sub chkResult_Click()
    If chkResult.Value = Checked Then
        lblResult.Caption = "PASS"
        lblResult.ForeColor = &H8000&
    Else
        lblResult.Caption = "FAIL"
        lblResult.ForeColor = vbRed
    End If
End Sub

Private Sub cmbProcess_change()
    Dim oPP As New clsPlasma_procesos
    Dim oPF As New clsPlasma_ficha
    Dim oPE As New clsPlasma_ensayos
    
    If cmbProcess.getTEXTO = "" Then
        txtDurezaReq = ""
    Else
'        oPP.Carga cmbProcess.getPK_SALIDA
'        oPF.Carga oPP.getBOND_COAT_FICHA_ID
'        If opTipo(0).Value = True Then ' Rockwell
'            txtDurezaReq = oPF.getMACRO_DUREZA_REQ
'            oPE.Carga oPF.getMACRO_DUREZA
'        Else ' Vicker
'            txtDurezaReq = oPF.getMICRO_DUREZA_REQ
'            oPE.Carga oPF.getMICRO_DUREZA
'        End If
'        Dim ounidad As New clsUnidades
'        ounidad.CARGAR oPE.getUNIDAD_ID
'        txtUnidades = ounidad.getNOMBRE
    End If
    Set oPP = Nothing
    Set oPF = Nothing
End Sub

Private Sub cmdImagen_Click()
    With frmCE_Imagenes
        .PK = PK
        .Show 1
    End With
End Sub

Private Sub cmdObservador_Click()

    Dim objfrm As New frmObservadorEnsayo

    objfrm.FORMULARIO_ORIGEN = 2 'Sellantes asociado al número 2
    objfrm.ES_CONTROL_EFICACIA = False
    objfrm.MUESTRA_ID = PK ' Id de la muestra
    objfrm.DETERMINACION_ENSAYO_ID = 0
    objfrm.SELLANTE_ID = txtID_SELLANTE
    objfrm.ENSAYO = lista.ListItems(lista.selectedItem.Index)
    
    If (UCase(lblCerrada) <> "CERRADA") Then
        objfrm.MUESTRA_CERRADA = False
    Else
        objfrm.MUESTRA_CERRADA = True
    End If

    objfrm.Show vbModal
    
    Set objfrm = Nothing

End Sub

Private Sub cmdok_Click()
    Dim oPRE As New clsPlasma_recepcion
   On Error GoTo cmdok_Click_Error

   On Error GoTo cmdok_Click_Error
    ' Validar equipos pendientes de verificacion
    Dim cont As Integer
    Dim oEV As New clsEquipoVerificacion
    Dim salidaVerificacion As String
    Dim salidaVerificacionAux As String
    For cont = 1 To listaEquipos.ListItems.Count
        salidaVerificacionAux = oEV.pendienteVerificacion(listaEquipos.ListItems(cont).Text, Date)
        If salidaVerificacionAux <> "" Then
            salidaVerificacion = salidaVerificacion & " - " & salidaVerificacionAux & vbNewLine
        End If
    Next
    If salidaVerificacion <> "" Then
        If MsgBox("ATENCIÓN : " & vbNewLine & salidaVerificacion & vbNewLine & " ¿Desea continuar?", vbExclamation + vbYesNo, App.Title) = vbNo Then
            Exit Sub
        End If
    End If
    ' Validaciones de campos
    Dim i As Integer
    Dim listaResultados() As String
    If txtDurezaResults <> "" Then
        listaResultados = Split(txtDurezaResults, "-")
        If UBound(listaResultados) <> 3 Then
            If MsgBox("ATENCIÓN : " & vbNewLine & " NO HA INTRODUCIDO 4 RESULTADOS " & vbNewLine & " ¿Desea continuar?", vbExclamation + vbYesNo, App.Title) = vbNo Then
                Exit Sub
            End If
        End If
    End If
    ' Validar rangos DUREZA
    If txtDurezaAverage <> "" Then
        If IsNumeric(txtDurezaAverage) Then
            If CInt(txtDurezaAverage) >= 55 And CInt(txtDurezaAverage) <= 85 Then
                chkResult.Value = Checked
            Else
                If CInt(txtDurezaAverage) < 55 Then
                    If MsgBox("El porcentaje de DUREZA es menor de 55. Se marcará el ensayo como FAIL. ¿Desea continuar?", vbCritical + vbYesNo, App.Title) = vbNo Then
                        Exit Sub
                    Else
                        chkResult.Value = Unchecked
                    End If
                End If
                If CInt(txtDurezaAverage) > 85 Then
                    If MsgBox("El porcentaje de DUREZA es mayor de 85. Se marcará el ensayo como FAIL. ¿Desea continuar?", vbCritical + vbYesNo, App.Title) = vbNo Then
                        Exit Sub
                    Else
                        chkResult.Value = Unchecked
                    End If
                End If
            End If
        End If
    End If
    ' Grabación de datos
    Me.MousePointer = 11
    With oPRE
        .setPROCESO_ID = cmbProcess.getPK_SALIDA
        .setCUSTOMER_ID = cmbCustomer.getPK_SALIDA
        .setSPECIMEN_ID = txtDatos(50)
        .setNTYPE = cmbnatype.getPK_SALIDA
        .setPN = txtDatos(51)
        .setPRODUCT_TYPE = txtDatos(52)
        .setMODULE_SN = txtDatos(53)
        .setSN = txtDatos(54)
        .setPRODUCT_SN = txtDatos(55)
        .setMP = 0
        .setMP_FECHA = "NULL"
        .setMP_USUARIO_ID = 0
        .setMP_PASS = 0
        .Modificar PK
        .ModificarResultado PK, chkResult.Value
        .informarControlSpecification PK
    End With
    Set oPRE = Nothing
    ' RESULTADOS
    Dim opd As New clsPlasma_dureza
    Dim res As String
    With opd
        .setMUESTRA_ID = PK
        .setIDENTIFICATION = "HARDNESS TEST (PNT IB 208)"
        .setDIMENSION = txtDurezaDimension
        .setESPESOR = txtDurezaEspesor
        .setREQUIREMENT = txtDurezaReq
        .setRESULT = txtDurezaResults
        .setAVERAGE = txtDurezaAverage
        If txtSD = "" Then
            .setSD = 0
        Else
            .setSD = Replace(txtSD, ",", ".")
        End If
        If txtPOR = "" Then
            .setPOR = 0
        Else
            .setPOR = Replace(txtPOR, ",", ".")
        End If
        .setPASS = chkResult.Value
        .Insertar
    End With
    
    Dim oPRH As New clsPlasma_resultados_historico
    oPRH.generar_dureza PK
    Set oPRH = Nothing
    
    Me.MousePointer = 0
    MsgBox "Datos almacenados correctamente.", vbInformation, App.Title
    If USUARIO.getPER_CIERRE = True Then
        Dim oMuestra As New clsMuestra
        oMuestra.CargaMuestra PK
        If oMuestra.getCERRADA = 0 Then
            If MsgBox("¿Desea cerrar la muestra?", vbQuestion + vbYesNo, App.Title) = vbYes Then
                oMuestra.Cerrar PK
            End If
        End If
    End If
    Unload Me
   On Error GoTo 0
   Exit Sub

cmdok_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure cmdok_Click of Formulario frmPlasma_ETR"
End Sub
Private Sub cmdVerificacion_Click()
    If listaEquipos.ListItems.Count > 0 Then
        Dim objfrm  As New frmEquipoEdicionVerificacion
        Dim oEquipo As New clsEquipos
        oEquipo.Carga listaEquipos.ListItems(listaEquipos.selectedItem.Index).Text
        Set objfrm.EQUIPO = oEquipo
        
        If listaEquipos.ListItems(listaEquipos.selectedItem.Index).SubItems(3) = 0 Then
            
            objfrm.TipoEdicion = Alta
            objfrm.idVerificadorInternoInicial = USUARIO.getID_EMPLEADO
            objfrm.FechaProximaInicial = Now
            objfrm.IdPeriodoInicial = ENUM_EQUIPOS_PERIODICIDAD.ENUM_EQUIPOS_PERIODICIDAD_ANTES_CADA_ENSAYO
            objfrm.IdTipoVerificacionIncial = 1
            
            objfrm.copiarUltimaVerificacionPeriodo = ENUM_EQUIPOS_PERIODICIDAD.ENUM_EQUIPOS_PERIODICIDAD_ANTES_CADA_ENSAYO
              
            objfrm.Show vbModal
            If objfrm.ID_VERIFICACION <> 0 Then
                listaEquipos.ListItems(listaEquipos.selectedItem.Index).SubItems(3) = objfrm.ID_VERIFICACION
            End If
            grabar_equipos
        Else
            objfrm.ID = listaEquipos.ListItems(listaEquipos.selectedItem.Index).SubItems(3)
            objfrm.TipoEdicion = visualizar
            objfrm.copiarUltimaVerificacionPeriodo = 0
            objfrm.Show vbModal
        End If
        
        Unload objfrm
        Set objfrm = Nothing
        Exit Sub
    End If
End Sub
Private Sub cmdAnadirEquipo_Click()
    If cmbEquipos.getPK_SALIDA <> 0 Then
        For i = 1 To listaEquipos.ListItems.Count
            If listaEquipos.ListItems(i) = cmbEquipos.getPK_SALIDA Then
                MsgBox "El equipo ya se encuentra en la lista.", vbExclamation, App.Title
                Exit Sub
            End If
        Next
        Dim oEquipo As New clsEquipos
        oEquipo.Carga_Datos_Basicos cmbEquipos.getPK_SALIDA
        With listaEquipos.ListItems.Add(, , oEquipo.getID_EQUIPO)
            .SubItems(1) = oEquipo.getNOMBRE
            .SubItems(2) = oEquipo.getSERIE
            .SubItems(3) = "0"
        End With
        listaEquipos.ListItems(listaEquipos.ListItems.Count).EnsureVisible
        cmbEquipos.limpiar
        grabar_equipos
    End If

End Sub

Private Sub cmdAnadirReactivo_Click()
    ' Interno (I)
    If cmbReactivos.getTEXTO <> "" Then
        Dim oBote As New clsBotes_ex
        Dim oTb As New clsTipos_bote_ex
        Dim oTR As New clsTipos_reactivo_ex
        oBote.CARGAR cmbReactivos.getPK_SALIDA
        oTb.CARGAR oBote.getTIPO_BOTE_EX_ID
        oTR.CARGAR oTb.getTIPO_REACTIVO_EX_ID
        With listaReactivos.ListItems.Add(, , oBote.getID_BOTE_EX)
            .SubItems(1) = oTR.getNOMBRE
            .SubItems(2) = Format(oBote.getFECHA_CADUCIDAD, "dd-mm-yyyy")
            .SubItems(3) = "E"
        End With
        listaReactivos.ListItems(listaReactivos.ListItems.Count).EnsureVisible
    End If
    ' Externo (E)
    If cmbReactivosInternos.getTEXTO <> "" Then
        Dim oRPR As New clsRpr_botes
        Dim oTRPR As New clsRPR_Tipos
        oRPR.Carga cmbReactivosInternos.getPK_SALIDA
        oTRPR.CARGAR oRPR.getTIPO_REACTIVO_PR_ID
        With listaReactivos.ListItems.Add(, , oRPR.getID_BOTE_PR)
            .SubItems(1) = oTRPR.getNOMBRE
            .SubItems(2) = Format(oRPR.getFECHA_CADUCIDAD, "dd-mm-yyyy")
            .SubItems(3) = "I"
        End With
        listaReactivos.ListItems(listaReactivos.ListItems.Count).EnsureVisible
    End If
    ' Limpiar Combos
    cmbReactivos.limpiar
    cmbReactivosInternos.limpiar
    grabar_reactivos
End Sub
Private Sub cmdEliminarEquipo_Click()
    If listaEquipos.ListItems.Count > 0 Then
        listaEquipos.ListItems.Remove listaEquipos.selectedItem.Index
        grabar_equipos
    End If
End Sub

Private Sub cmdEliminarReactivo_Click()
    If listaReactivos.ListItems.Count > 0 Then
        listaReactivos.ListItems.Remove listaReactivos.selectedItem.Index
        cmbReactivos.limpiar
        cmbReactivosInternos.limpiar
    End If
    grabar_reactivos
End Sub
Private Sub cmdSalir_Click()
'    grabar_equipos
    Dim oMuestra As New clsMuestra
    oMuestra.comprobar_cierre (PK)
    Unload Me
End Sub
Private Sub Form_Load()
    log (Me.Name)
    cargar_botones Me
    cabecera
    cargar_combos
    permisos
    inicializar_grid
    If PK > 0 Then
        cargar_muestra
    End If
End Sub
Private Sub permisos()
    ' Permiso para modificar la vida
    Dim op As New clsParametros
    Dim s() As String
    Dim i As Integer
    op.Carga parametros.PARAM_USUARIOS_MODIFICAN_EQUIPOS_MUESTRA_CERRADA, ""
    If op.getVALOR <> "" Then
        s = Split(op.getVALOR, ",")
        For i = LBound(s) To UBound(s)
            If USUARIO.getID_EMPLEADO = CInt(s(i)) Then
                chkModificar.Value = Checked
                Exit For
            End If
        Next
    End If
    Set op = Nothing
End Sub

Private Sub cabecera()
    With listaEquipos.ColumnHeaders
        .Add , , "NºEquipo", 800, lvwColumnLeft
        .Add , , "Nombre", 3200, lvwColumnLeft
        .Add , , "NºSerie", 1200, lvwColumnCenter
        .Add , , "Verificación", 1, lvwColumnCenter
    End With
    With listaReactivos.ColumnHeaders
        .Add , , "Número", 800, lvwColumnLeft
        .Add , , "Reactivo", 3200, lvwColumnLeft
        .Add , , "Caducidad", 1200, lvwColumnCenter
        .Add , , "TIPO", 0, lvwColumnCenter ' (I-E) Interno o externo
    End With
End Sub
Private Sub cargar_muestra()
    'Titulo
    Dim oMuestra As New clsMuestra
   On Error GoTo cargar_muestra_Error

    oMuestra.CargaMuestra (PK)
    ' Duplicada
    If oMuestra.getANALISIS_DUPLICADO = 1 Then
        chkDuplicada.Value = Checked
    End If
    lbltitulo = "Registro resultados : " & Trim(str(oMuestra.getID_GENERAL)) & " (" & oMuestra.CodigoParticular(PK) & ")"
    Me.Caption = lbltitulo
    'Equipos
    cargar_equipos PK
    cargar_reactivos PK
    ' Cargar datos de recepción
    Dim oPlasmaRecepcion As New clsPlasma_recepcion
    With oPlasmaRecepcion
        If .Carga(PK) Then
            cmbProcess.MostrarElemento .getPROCESO_ID
            cmbCustomer.MostrarElemento .getCUSTOMER_ID
            cmbnatype.MostrarElemento .getNTYPE
            txtDatos(50) = .getSPECIMEN_ID
            txtDatos(51) = .getPN
            txtDatos(52) = .getPRODUCT_TYPE
            txtDatos(53) = .getMODULE_SN
            txtDatos(54) = .getSN
            txtDatos(55) = .getPRODUCT_SN
            chkResult = .getRESULT
        End If
    End With
    ' Resultados
'    Dim opd As New clsPlasma_dureza
'    If opd.Carga(PK) = True Then
'        txtDurezaResults = opd.getRESULT
'        txtDurezaReq = opd.getREQUIREMENT
'        txtDurezaAverage = opd.getAVERAGE
'        txtDurezaDimension = opd.getDIMENSION
'        txtDurezaEspesor = opd.getESPESOR
'        txtSD = opd.getSD
'        txtPOR = opd.getPOR
'    End If
'    If txtDurezaReq = "" Then
'        Dim oPP As New clsPlasma_procesos
'        Dim oPF As New clsPlasma_ficha
'        oPP.Carga oPlasmaRecepcion.getPROCESO_ID
'        oPF.Carga oPP.getBOND_COAT_FICHA_ID
'
'        txtDurezaReq = oPF.getSHOREA_REQ
'    End If
'    If txtDurezaDimension = "" Then
'        txtDurezaDimension = "ATTACHED REPAIR ORDER"
'    End If
    
    Set oPlasmaRecepcion = Nothing
    Set opd = Nothing
    proteger_campos oMuestra.getCERRADA

   On Error GoTo 0
   Exit Sub

cargar_muestra_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure cargar_muestra of Formulario frmPlasma_ETR"
End Sub

Private Sub listaEquipos_DblClick()
    If listaEquipos.ListItems.Count > 0 Then
        frmEquipoEdicion.PK = CLng(listaEquipos.ListItems(listaEquipos.selectedItem.Index).Text)
        frmEquipoEdicion.Show 1
    End If
End Sub

Private Sub listaEquipos_ItemCheck(ByVal Item As MSComctlLib.ListItem)
    grabar_equipos
End Sub
Private Sub txtdatos_GotFocus(Index As Integer)
    txtDatos(Index).BackColor = vbYellow
    txtDatos(Index).SelStart = 0
    txtDatos(Index).SelLength = Len(txtDatos(Index))
End Sub

Private Sub txtdatos_LostFocus(Index As Integer)
    txtDatos(Index).BackColor = vbWhite
    If Index = 8 Or Index = 9 Then
        If Trim(txtDatos(Index)) <> "" Then
            If Right(txtDatos(Index), 2) <> "ºC" Then
                txtDatos(Index) = txtDatos(Index) & " ºC"
            End If
        End If
    End If
End Sub

Private Sub txtvalor_GotFocus()
    txtValor.BackColor = vbYellow
    txtValor.SelStart = 0
    txtValor.SelLength = Len(txtValor)
End Sub
Private Sub txtvalor_LostFocus()
    txtValor.BackColor = vbWhite
End Sub
Private Sub proteger_campos(CERRADA As Integer)
    If (CERRADA = 1 Or CERRADA = 3) And chkModificar.Value = Unchecked Then
        cmdEliminarReactivo.Enabled = False
        cmdAnadirReactivo.Enabled = False
        cmdEliminarEquipo.Enabled = False
        cmdAnadirEquipo.Enabled = False
        cmbEquipos.desactivar
        cmbReactivos.desactivar
        cmbReactivosInternos.desactivar
        cmbProcess.desactivar
        cmbCustomer.desactivar
        txtDatos(50).Enabled = False
        txtDatos(51).Enabled = False
        txtDatos(52).Enabled = False
        txtDatos(53).Enabled = False
        txtDatos(54).Enabled = False
        txtDatos(55).Enabled = False
        chkResult.Enabled = False
        cmdok.visible = False
    Else
        cmdEliminarReactivo.Enabled = True
        cmdAnadirReactivo.Enabled = True
        cmdEliminarEquipo.Enabled = True
        cmdAnadirEquipo.Enabled = True
        cmbEquipos.activar
        cmbReactivos.activar
        cmbReactivosInternos.activar
        cmbProcess.activar
        cmbCustomer.activar
        txtDatos(50).Enabled = True
        txtDatos(51).Enabled = True
        txtDatos(52).Enabled = True
        txtDatos(53).Enabled = True
        txtDatos(54).Enabled = True
        txtDatos(55).Enabled = True
        chkResult.Enabled = True
        cmdok.visible = True
    End If
    Select Case CERRADA
        Case 0
            lblCerrada = "ABIERTA"
        Case 1
            lblCerrada = "CERRADA"
            lblCerrada.BackColor = vbRed
        Case 2
            lblCerrada = "PTE. CIERRE"
        Case 3
            lblCerrada = "C.SIN INFORME"
    End Select
End Sub

Private Sub cargar_equipos(muestra As Long)
    Dim oPE As New clsPlasma_equipos
    Dim rs As ADODB.Recordset
    Set rs = oPE.Listado(muestra)
    listaEquipos.ListItems.Clear
    If rs.RecordCount > 0 Then
        Do
            With listaEquipos.ListItems.Add(, , rs(0))
                .SubItems(1) = rs(1)
                .SubItems(2) = rs(2)
                .SubItems(3) = rs(5) ' VERIFICACION
            End With
            If rs("EN_INFORME") = 1 Then
                listaEquipos.ListItems(listaEquipos.ListItems.Count).Checked = True
            Else
                listaEquipos.ListItems(listaEquipos.ListItems.Count).Checked = False
            End If
            rs.MoveNext
        Loop Until rs.EOF
    End If
    Set oPE = Nothing
    
End Sub

Private Sub cargar_reactivos(muestra As Long)
    Dim oPR As New clsPlasma_Reactivos
    Dim oReactivo As New clsBotes_ex
    Dim oTb As New clsTipos_bote_ex
    Dim oTR As New clsTipos_reactivo_ex
    
    Dim oRPR As New clsRpr_botes
    Dim oTRPR As New clsRPR_Tipos
    Dim rs As ADODB.Recordset
    Set rs = oPR.Listado(muestra)
    If rs.RecordCount > 0 Then
        Do
            If rs(1) = "E" Then
               oReactivo.CARGAR CLng(rs(0))
               oTb.CARGAR oReactivo.getTIPO_BOTE_EX_ID
               oTR.CARGAR oTb.getTIPO_REACTIVO_EX_ID
               With listaReactivos.ListItems.Add(, , rs(0))
                  .SubItems(1) = oTR.getNOMBRE
                  .SubItems(2) = Format(oReactivo.getFECHA_CADUCIDAD, "DD-MM-YYYY")
                  .SubItems(3) = "E"
               End With
            Else
                oRPR.Carga CLng(rs(0))
                oTRPR.CARGAR oRPR.getTIPO_REACTIVO_PR_ID
                With listaReactivos.ListItems.Add(, , rs(0))
                    .SubItems(1) = oTRPR.getCODIGO & "-" & Format(oRPR.getNUMERO, "000") & " " & oTRPR.getNOMBRE
                    .SubItems(2) = Format(oRPR.getFECHA_CADUCIDAD, "DD-MM-YYYY")
                    .SubItems(3) = "I"
                End With
            End If
            rs.MoveNext
        Loop Until rs.EOF
    End If
End Sub

Private Sub cargar_combos()
    llenar_combo cmbProcess, New clsPlasma_procesos, 0, frmPlasma_Procesos_Detalle, ""
    Dim oDeco As New clsDecodificadora
    oDeco.cargar_mi_combo cmbCustomer, DECODIFICADORA.DECODIFICADORA_PLASMA_CLIENTES_INTERNOS
    oDeco.cargar_mi_combo cmbnatype, DECODIFICADORA.DECODIFICADORA_PLASMA_NUMBER_AND_TYPE
    
    llenar_combo cmbEquipos, New clsEquipos, 0, frmEquipoEdicion, ""
    llenar_combo cmbReactivos, New clsBotes_ex, 0, Me, " AND ABIERTO = 1 AND FINALIZADO = 0 "
    llenar_combo cmbReactivosInternos, New clsRpr_botes, 0, Me, " AND ISNULL(fecha_fin)"
End Sub
Private Sub grabar_equipos()
    Dim Equipos As String
    Dim oPE As New clsPlasma_equipos
    oPE.Eliminar PK
    Dim i As Integer
    For i = 1 To listaEquipos.ListItems.Count
        Equipos = Equipos & listaEquipos.ListItems(i).Text & ";"
        With oPE
            .setMUESTRA_ID = PK
            .setORDEN = i
            .setEQUIPO_ID = listaEquipos.ListItems(i).Text
            .setVERIFICACION_ID = listaEquipos.ListItems(i).SubItems(3)
            .setEN_INFORME = Abs(listaEquipos.ListItems(i).Checked)
            .Insertar
        End With
    Next
    ' Usos de los equipos
    Dim oEU As New clsEq_usos
    oEU.Eliminar PK, 0
    For i = 1 To listaEquipos.ListItems.Count
      With oEU
          .setMUESTRA_ID = PK
          .setEQUIPO_ID = listaEquipos.ListItems(i).Text
          .setDETERMINACION_ID = 0
          .setUSOS = 1
          .Insertar
      End With
    Next
    Set oEU = Nothing
End Sub
Private Sub grabar_reactivos()
    Dim oPR As New clsPlasma_Reactivos
    oPR.Eliminar PK
    Dim i As Integer
    For i = 1 To listaReactivos.ListItems.Count
        With oPR
            .setMUESTRA_ID = PK
            .setBOTE_EX_ID = listaReactivos.ListItems(i).Text
            .setTIPO = listaReactivos.ListItems(i).SubItems(3)
            .setORDEN = i
            .Insertar
        End With
    Next
    Set oPR = Nothing
End Sub

Private Sub inicializar_grid()
   On Error GoTo inicializar_grid_Error

    xTRACCION.Clear
    xTRACCION.ReDim 0, filasGrid, 0, ColGrid
    xTRACCION.Clear
    ' Configuración inicial del grid
'    xTRACCION(0, ColsGrid.IDENTIFICATION) = "P1"
'    xTRACCION(1, ColsGrid.IDENTIFICATION) = "P2"
'    xTRACCION(2, ColsGrid.IDENTIFICATION) = "P3"
'    xTRACCION(3, ColsGrid.IDENTIFICATION) = "B (Blank Strength)"
'
'    xTRACCION(0, ColsGrid.LOCATION) = "Intracoating: -- % / In epoxy : -- %"
'    xTRACCION(1, ColsGrid.LOCATION) = "Intracoating: -- % / In epoxy : -- %"
'    xTRACCION(2, ColsGrid.LOCATION) = "Intracoating: -- % / In epoxy : -- %"
'
'    txtRoom = "-- ºC / -- %Hr"
'    ' Cargar de los datos de la muestra
'    Dim rs As ADODB.Recordset
'    Dim oPTP As New clsPlasma_traccion_p
'    Set rs = oPTP.Listado(MUESTRA_ID, tipo)
'    If rs.RecordCount > 0 Then
'        Do
'            xTRACCION(rs("ORDEN"), ColsGrid.DIAMETER) = CStr(rs("DIAMETER"))
'            xTRACCION(rs("ORDEN"), ColsGrid.AREA) = CStr(rs("AREA"))
'            xTRACCION(rs("ORDEN"), ColsGrid.LOADP) = CStr(rs("LOADP"))
'            xTRACCION(rs("ORDEN"), ColsGrid.TENSILE) = CStr(rs("TENSILE"))
'            xTRACCION(rs("ORDEN"), ColsGrid.LOCATION) = CStr(rs("LOCATION"))
'
'            rs.MoveNext
'        Loop Until rs.EOF
'    End If
'    Set rs = Nothing
    Set grid.Array = xTRACCION
    grid.Refresh
   On Error GoTo 0
   Exit Sub

inicializar_grid_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure inicializar_grid of Formulario frmCE_Recepcion_Nuevo2"
End Sub
Private Sub grid_KeyPress(KeyAscii As Integer)
    If KeyAscii = 46 Then
         KeyAscii = 44
    End If
    If KeyAscii = 13 Then
        calcularResultados
    End If
End Sub

Private Sub calcularResultados()
    Dim i As Integer
    ' Calculo de Area y Tensile
   On Error GoTo calcularResultados_Error

    For i = 0 To filasGrid
        If Not IsEmpty(xTRACCION(i, ColsGrid.IDENTIFICATION)) Then
            Dim area As Single
            
            If Trim(xTRACCION(i, ColsGrid.LADO1)) <> "" And Trim(xTRACCION(i, ColsGrid.LADO2)) <> "" Then
                area = CSng(Trim(xTRACCION(i, ColsGrid.LADO1))) * CSng(Trim(xTRACCION(i, ColsGrid.LADO2)))
                xTRACCION(i, ColsGrid.area) = area
                xTRACCION(i, ColsGrid.RESULTADO_PSI) = CInt(CSng(xTRACCION(i, ColsGrid.VALOR_OBTENIDO)) / area)
                xTRACCION(i, ColsGrid.RESULTADO_MPA) = (CSng(xTRACCION(i, ColsGrid.VALOR_OBTENIDO)) / area) / 145
            End If
        End If
    Next
    grid.Refresh

   On Error GoTo 0
   Exit Sub

calcularResultados_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure calcularResultados of Formulario frmPlasma_ETR"
End Sub

