VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{CEA3B7F6-2847-4E5E-A551-DB7A62489D44}#46.0#0"; "miCombo.ocx"
Object = "{77DDF307-D82B-4757-8B3A-106EC9D75D4B}#8.0#0"; "tdbg8.ocx"
Begin VB.Form frmRPR_Reactivo 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Reactivo Propio/Suministro"
   ClientHeight    =   9060
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   12975
   Icon            =   "frmRPR_Reactivo.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   9060
   ScaleWidth      =   12975
   StartUpPosition =   2  'CenterScreen
   Begin TrueDBGrid80.TDBDropDown tUnidades 
      Height          =   2730
      Left            =   10530
      TabIndex        =   29
      Top             =   5310
      Width           =   2370
      _ExtentX        =   4180
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
   Begin VB.Frame Frame3 
      BackColor       =   &H00C0C0C0&
      Height          =   1050
      Left            =   135
      TabIndex        =   24
      Top             =   7965
      Width           =   10140
      Begin VB.CommandButton cmdAdd 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         Caption         =   "Añadir"
         Height          =   825
         Left            =   8055
         Picture         =   "frmRPR_Reactivo.frx":030A
         Style           =   1  'Graphical
         TabIndex        =   28
         Top             =   135
         Width           =   960
      End
      Begin VB.CommandButton cmdReset 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         Caption         =   "Eliminar"
         Height          =   825
         Left            =   9090
         Picture         =   "frmRPR_Reactivo.frx":0BD4
         Style           =   1  'Graphical
         TabIndex        =   27
         Top             =   135
         Width           =   960
      End
      Begin pryCombo.miCombo cmbReactivos 
         Height          =   330
         Left            =   1530
         TabIndex        =   25
         Top             =   195
         Width           =   6405
         _ExtentX        =   11298
         _ExtentY        =   582
      End
      Begin pryCombo.miCombo cmbPropio 
         Height          =   330
         Left            =   1515
         TabIndex        =   32
         Top             =   600
         Width           =   6405
         _ExtentX        =   11298
         _ExtentY        =   582
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Reactivo Propio"
         Height          =   195
         Index           =   8
         Left            =   120
         TabIndex        =   33
         Top             =   645
         Width           =   1140
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Reactivo Externo"
         Height          =   195
         Index           =   6
         Left            =   135
         TabIndex        =   26
         Top             =   240
         Width           =   1230
      End
   End
   Begin VB.CommandButton cmdok 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Aceptar"
      Height          =   870
      Left            =   10740
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   8100
      Width           =   1050
   End
   Begin VB.CommandButton cmdcancel 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Cancelar"
      Height          =   870
      Left            =   11835
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   8100
      Width           =   1050
   End
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      ForeColor       =   &H80000008&
      Height          =   4050
      Left            =   45
      TabIndex        =   10
      Top             =   675
      Width           =   12810
      Begin VB.Frame Frame2 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Tipo"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   645
         Left            =   8460
         TabIndex        =   21
         Top             =   135
         Width           =   4245
         Begin VB.OptionButton optTipo 
            BackColor       =   &H00C0C0C0&
            Caption         =   "Suministro"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   240
            Index           =   2
            Left            =   2565
            TabIndex        =   23
            Top             =   270
            Width           =   1410
         End
         Begin VB.OptionButton optTipo 
            BackColor       =   &H00C0C0C0&
            Caption         =   "Reactivo propio"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   240
            Index           =   1
            Left            =   360
            TabIndex        =   22
            Top             =   270
            Value           =   -1  'True
            Width           =   1995
         End
      End
      Begin VB.TextBox txtDatos 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   360
         Index           =   6
         Left            =   1680
         TabIndex        =   5
         Top             =   3240
         Width           =   1530
      End
      Begin VB.TextBox txtDatos 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   900
         Index           =   4
         Left            =   1680
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   4
         Top             =   2295
         Width           =   10950
      End
      Begin VB.TextBox txtDatos 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   360
         Index           =   5
         Left            =   4680
         TabIndex        =   6
         Top             =   3240
         Width           =   3150
      End
      Begin VB.TextBox txtDatos 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
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
         Index           =   0
         Left            =   1680
         TabIndex        =   0
         Top             =   450
         Width           =   2445
      End
      Begin VB.TextBox txtDatos 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   345
         Index           =   1
         Left            =   1680
         MaxLength       =   250
         TabIndex        =   1
         Top             =   825
         Width           =   10950
      End
      Begin VB.TextBox txtDatos 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   330
         Index           =   2
         Left            =   1680
         MaxLength       =   250
         TabIndex        =   2
         Top             =   1200
         Width           =   10950
      End
      Begin VB.TextBox txtDatos 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   705
         Index           =   3
         Left            =   1680
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   3
         Top             =   1560
         Width           =   10950
      End
      Begin MSDataListLib.DataCombo cmbcad 
         Height          =   315
         Left            =   9240
         TabIndex        =   7
         Top             =   3255
         Width           =   3390
         _ExtentX        =   5980
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
      Begin MSDataListLib.DataCombo cmbCentro 
         Bindings        =   "frmRPR_Reactivo.frx":149E
         Height          =   315
         Left            =   9240
         TabIndex        =   34
         Top             =   3645
         Width           =   3390
         _ExtentX        =   5980
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
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Centro"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   22
         Left            =   8145
         TabIndex        =   35
         Top             =   3690
         Width           =   645
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Volumen Total"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   7
         Left            =   180
         TabIndex        =   18
         Top             =   3315
         Width           =   1335
      End
      Begin VB.Label lblCampos 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Caducidad"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   13
         Left            =   8145
         TabIndex        =   17
         Top             =   3330
         Width           =   1005
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Protocolo"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   5
         Left            =   180
         TabIndex        =   16
         Top             =   2745
         Width           =   885
      End
      Begin VB.Label lblCampos 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "P.Referencia "
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   4
         Left            =   3345
         TabIndex        =   15
         Top             =   3330
         Width           =   1245
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Codigo"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   0
         Left            =   180
         TabIndex        =   14
         Top             =   495
         Width           =   660
      End
      Begin VB.Label lblCampos 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Nombre"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   1
         Left            =   180
         TabIndex        =   13
         Top             =   900
         Width           =   765
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Almacenaje"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   2
         Left            =   180
         TabIndex        =   12
         Top             =   1305
         Width           =   1080
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Equipos"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   3
         Left            =   180
         TabIndex        =   11
         Top             =   1845
         Width           =   750
      End
   End
   Begin TrueDBGrid80.TDBGrid gridComponentes 
      Height          =   2820
      Left            =   90
      TabIndex        =   20
      Top             =   5040
      Width           =   12795
      _ExtentX        =   22569
      _ExtentY        =   4974
      _LayoutType     =   4
      _RowHeight      =   -2147483647
      _WasPersistedAsPixels=   0
      Columns(0)._VlistStyle=   0
      Columns(0)._MaxComboItems=   5
      Columns(0).Caption=   "Reactivo"
      Columns(0).DataField=   ""
      Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(1)._VlistStyle=   0
      Columns(1)._MaxComboItems=   5
      Columns(1).Caption=   "P. Referencia"
      Columns(1).DataField=   ""
      Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(2)._VlistStyle=   0
      Columns(2)._MaxComboItems=   5
      Columns(2).Caption=   "Cantidad"
      Columns(2).DataField=   ""
      Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(3)._VlistStyle=   0
      Columns(3)._MaxComboItems=   5
      Columns(3).Caption=   "Unidad"
      Columns(3).DataField=   ""
      Columns(3).DropDown=   "tUnidades"
      Columns(3).DropDown.vt=   8
      Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(4)._VlistStyle=   0
      Columns(4)._MaxComboItems=   5
      Columns(4).Caption=   "ID"
      Columns(4).DataField=   ""
      Columns(4).DropDown=   "tResponsables"
      Columns(4).DropDown.vt=   8
      Columns(4)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(5)._VlistStyle=   0
      Columns(5)._MaxComboItems=   5
      Columns(5).Caption=   "ID_UNIDAD"
      Columns(5).DataField=   ""
      Columns(5)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(6)._VlistStyle=   0
      Columns(6)._MaxComboItems=   5
      Columns(6).Caption=   "TIPO"
      Columns(6).DataField=   ""
      Columns(6)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns.Count   =   7
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
      Splits(0)._ColumnProps(0)=   "Columns.Count=7"
      Splits(0)._ColumnProps(1)=   "Column(0).Width=10689"
      Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
      Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=10610"
      Splits(0)._ColumnProps(4)=   "Column(0)._EditAlways=0"
      Splits(0)._ColumnProps(5)=   "Column(0)._ColStyle=8193"
      Splits(0)._ColumnProps(6)=   "Column(0).Order=1"
      Splits(0)._ColumnProps(7)=   "Column(1).Width=4974"
      Splits(0)._ColumnProps(8)=   "Column(1).DividerColor=0"
      Splits(0)._ColumnProps(9)=   "Column(1)._WidthInPix=4895"
      Splits(0)._ColumnProps(10)=   "Column(1)._EditAlways=0"
      Splits(0)._ColumnProps(11)=   "Column(1)._ColStyle=1"
      Splits(0)._ColumnProps(12)=   "Column(1).Order=2"
      Splits(0)._ColumnProps(13)=   "Column(2).Width=2725"
      Splits(0)._ColumnProps(14)=   "Column(2).DividerColor=0"
      Splits(0)._ColumnProps(15)=   "Column(2)._WidthInPix=2646"
      Splits(0)._ColumnProps(16)=   "Column(2)._EditAlways=0"
      Splits(0)._ColumnProps(17)=   "Column(2)._ColStyle=1"
      Splits(0)._ColumnProps(18)=   "Column(2).Order=3"
      Splits(0)._ColumnProps(19)=   "Column(3).Width=2725"
      Splits(0)._ColumnProps(20)=   "Column(3).DividerColor=0"
      Splits(0)._ColumnProps(21)=   "Column(3)._WidthInPix=2646"
      Splits(0)._ColumnProps(22)=   "Column(3)._EditAlways=0"
      Splits(0)._ColumnProps(23)=   "Column(3).Order=4"
      Splits(0)._ColumnProps(24)=   "Column(3).AutoDropDown=1"
      Splits(0)._ColumnProps(25)=   "Column(4).Width=1482"
      Splits(0)._ColumnProps(26)=   "Column(4).DividerColor=0"
      Splits(0)._ColumnProps(27)=   "Column(4)._WidthInPix=1402"
      Splits(0)._ColumnProps(28)=   "Column(4)._EditAlways=0"
      Splits(0)._ColumnProps(29)=   "Column(4)._ColStyle=1"
      Splits(0)._ColumnProps(30)=   "Column(4).Visible=0"
      Splits(0)._ColumnProps(31)=   "Column(4).Order=5"
      Splits(0)._ColumnProps(32)=   "Column(4).DropDownList=1"
      Splits(0)._ColumnProps(33)=   "Column(5).Width=2725"
      Splits(0)._ColumnProps(34)=   "Column(5).DividerColor=0"
      Splits(0)._ColumnProps(35)=   "Column(5)._WidthInPix=2646"
      Splits(0)._ColumnProps(36)=   "Column(5)._EditAlways=0"
      Splits(0)._ColumnProps(37)=   "Column(5).Visible=0"
      Splits(0)._ColumnProps(38)=   "Column(5).Order=6"
      Splits(0)._ColumnProps(39)=   "Column(6).Width=873"
      Splits(0)._ColumnProps(40)=   "Column(6).DividerColor=0"
      Splits(0)._ColumnProps(41)=   "Column(6)._WidthInPix=794"
      Splits(0)._ColumnProps(42)=   "Column(6)._EditAlways=0"
      Splits(0)._ColumnProps(43)=   "Column(6).Order=7"
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
      _StyleDefs(36)  =   "Splits(0).Columns(0).Style:id=24,.parent=11,.alignment=2,.locked=-1"
      _StyleDefs(37)  =   "Splits(0).Columns(0).HeadingStyle:id=21,.parent=12"
      _StyleDefs(38)  =   "Splits(0).Columns(0).FooterStyle:id=22,.parent=13"
      _StyleDefs(39)  =   "Splits(0).Columns(0).EditorStyle:id=23,.parent=15"
      _StyleDefs(40)  =   "Splits(0).Columns(1).Style:id=32,.parent=11,.alignment=2"
      _StyleDefs(41)  =   "Splits(0).Columns(1).HeadingStyle:id=29,.parent=12"
      _StyleDefs(42)  =   "Splits(0).Columns(1).FooterStyle:id=30,.parent=13"
      _StyleDefs(43)  =   "Splits(0).Columns(1).EditorStyle:id=31,.parent=15"
      _StyleDefs(44)  =   "Splits(0).Columns(2).Style:id=54,.parent=11,.alignment=2"
      _StyleDefs(45)  =   "Splits(0).Columns(2).HeadingStyle:id=51,.parent=12"
      _StyleDefs(46)  =   "Splits(0).Columns(2).FooterStyle:id=52,.parent=13"
      _StyleDefs(47)  =   "Splits(0).Columns(2).EditorStyle:id=53,.parent=15"
      _StyleDefs(48)  =   "Splits(0).Columns(3).Style:id=28,.parent=11"
      _StyleDefs(49)  =   "Splits(0).Columns(3).HeadingStyle:id=25,.parent=12"
      _StyleDefs(50)  =   "Splits(0).Columns(3).FooterStyle:id=26,.parent=13"
      _StyleDefs(51)  =   "Splits(0).Columns(3).EditorStyle:id=27,.parent=15"
      _StyleDefs(52)  =   "Splits(0).Columns(4).Style:id=36,.parent=11,.alignment=2"
      _StyleDefs(53)  =   "Splits(0).Columns(4).HeadingStyle:id=33,.parent=12"
      _StyleDefs(54)  =   "Splits(0).Columns(4).FooterStyle:id=34,.parent=13"
      _StyleDefs(55)  =   "Splits(0).Columns(4).EditorStyle:id=35,.parent=15"
      _StyleDefs(56)  =   "Splits(0).Columns(5).Style:id=58,.parent=11"
      _StyleDefs(57)  =   "Splits(0).Columns(5).HeadingStyle:id=55,.parent=12"
      _StyleDefs(58)  =   "Splits(0).Columns(5).FooterStyle:id=56,.parent=13"
      _StyleDefs(59)  =   "Splits(0).Columns(5).EditorStyle:id=57,.parent=15"
      _StyleDefs(60)  =   "Splits(0).Columns(6).Style:id=62,.parent=11"
      _StyleDefs(61)  =   "Splits(0).Columns(6).HeadingStyle:id=59,.parent=12"
      _StyleDefs(62)  =   "Splits(0).Columns(6).FooterStyle:id=60,.parent=13"
      _StyleDefs(63)  =   "Splits(0).Columns(6).EditorStyle:id=61,.parent=15"
      _StyleDefs(64)  =   "Named:id=37:Normal"
      _StyleDefs(65)  =   ":id=37,.parent=0,.alignment=3,.bgcolor=&H80000018&,.bgpicMode=2,.bold=0"
      _StyleDefs(66)  =   ":id=37,.fontsize=825,.italic=0,.underline=0,.strikethrough=0,.charset=0"
      _StyleDefs(67)  =   ":id=37,.fontname=MS Sans Serif"
      _StyleDefs(68)  =   "Named:id=38:Heading"
      _StyleDefs(69)  =   ":id=38,.parent=37,.valignment=2,.bgcolor=&H80000016&,.fgcolor=&H80000012&"
      _StyleDefs(70)  =   ":id=38,.wraptext=-1,.bold=0,.fontsize=825,.italic=0,.underline=0"
      _StyleDefs(71)  =   ":id=38,.strikethrough=0,.charset=0"
      _StyleDefs(72)  =   ":id=38,.fontname=MS Sans Serif"
      _StyleDefs(73)  =   "Named:id=39:Footing"
      _StyleDefs(74)  =   ":id=39,.parent=37,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(75)  =   "Named:id=40:Selected"
      _StyleDefs(76)  =   ":id=40,.parent=37,.bgcolor=&H8080FF&,.fgcolor=&H0&,.bold=0,.fontsize=825"
      _StyleDefs(77)  =   ":id=40,.italic=0,.underline=0,.strikethrough=0,.charset=0"
      _StyleDefs(78)  =   ":id=40,.fontname=MS Sans Serif"
      _StyleDefs(79)  =   "Named:id=41:Caption"
      _StyleDefs(80)  =   ":id=41,.parent=38,.alignment=2"
      _StyleDefs(81)  =   "Named:id=42:HighlightRow"
      _StyleDefs(82)  =   ":id=42,.parent=37,.bgcolor=&H80000008&,.fgcolor=&H80000005&"
      _StyleDefs(83)  =   "Named:id=43:EvenRow"
      _StyleDefs(84)  =   ":id=43,.parent=37,.bgcolor=&HFFFFFF&,.wraptext=-1"
      _StyleDefs(85)  =   "Named:id=44:OddRow"
      _StyleDefs(86)  =   ":id=44,.parent=37,.bgcolor=&HD5ECF9&"
      _StyleDefs(87)  =   "Named:id=47:RecordSelector"
      _StyleDefs(88)  =   ":id=47,.parent=38"
      _StyleDefs(89)  =   "Named:id=50:FilterBar"
      _StyleDefs(90)  =   ":id=50,.parent=37"
   End
   Begin VB.Label lbldeter 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0FF&
      Caption         =   "Componentes"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   300
      Left            =   90
      TabIndex        =   19
      Top             =   4770
      Width           =   12735
   End
   Begin VB.Label lbltitulo 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "Detalle de preparaciónd de reactivo Propio / Suministro"
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
      Index           =   0
      Left            =   90
      TabIndex        =   31
      Top             =   90
      Width           =   5805
   End
   Begin VB.Image imagen 
      Height          =   480
      Left            =   12330
      Picture         =   "frmRPR_Reactivo.frx":14E4
      Top             =   90
      Width           =   480
   End
   Begin VB.Label lbltitulo 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "Detalle de preparaciónd de reactivo Propio / Suministro"
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   1
      Left            =   90
      TabIndex        =   30
      Top             =   405
      Width           =   3915
   End
   Begin VB.Shape fondo 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   690
      Left            =   0
      Top             =   0
      Width           =   12915
   End
End
Attribute VB_Name = "frmRPR_Reactivo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim xP As New XArrayDB
Dim xUnidades As New XArrayDB
Const filasP As Integer = 50
Const ColP As Integer = 6
Private Enum ColsP
    Reactivo = 0
    P_REFERENCIA = 1
    cantidad = 2
    Unidad = 3
    ID = 4
    id_unidad = 5
    tipo = 6
End Enum
Const ColUnidades As Integer = 1
Private Enum ColsUnidades
    Unidad = 0
    ID = 1
End Enum
Private Sub cmdAdd_Click()
    If cmbReactivos.getTEXTO <> "" Then
        If cmbReactivos.getPK_SALIDA > 0 Then
           Call anadir_reactivo(cmbReactivos.getTEXTO, cmbReactivos.getPK_SALIDA, "", "", "", "", "E")
           gridComponentes.Refresh
        End If
    ElseIf cmbPropio.getTEXTO <> "" Then
        If cmbPropio.getPK_SALIDA > 0 Then
           Call anadir_reactivo(cmbPropio.getTEXTO, cmbPropio.getPK_SALIDA, "", "", "", "", "I")
           gridComponentes.Refresh
        End If
    End If
End Sub

Private Sub cmdcancel_Click()
    Unload Me
End Sub

'E0150-I
Private Sub cmdok_Click()
    Dim auxiliar As Long
    Dim strTipo As String
    
   On Error GoTo cmdok_Click_Error

    If datos_correctos Then
        ' Reactivo/suministro
      auxiliar = greactivopr
      Dim oTipos_reactivo_pr As New clsRPR_Tipos
      With oTipos_reactivo_pr
            .setCODIGO = txtDatos(0)
            .setCENTRO_ID = cmbCentro.BoundText
            If optTipo(1).value = True Then ' Si es un reactivo propio
                .setTIPO = 1 ' Es un reactivo propio
                strTipo = "Reactivo Propio"
            Else
                .setTIPO = 2 ' Es un suministro
                strTipo = "Suministro"
            End If
            .setNOMBRE = txtDatos(1)
            If txtDatos(2) = "" Then
            .setALMACENAMIENTO = "No existen condiciones especiales."
          Else
              .setALMACENAMIENTO = txtDatos(2)
          End If
            .setEQUIPOS = txtDatos(3)
            .setPROTOCOLO = txtDatos(4)
            .setPROC_REFERENCIA = txtDatos(5)
            .setCANTIDAD = txtDatos(6)
          .setTIPO_CADUCIDAD_ID = cmbcad.BoundText
      End With
      If greactivopr = 0 Then
            If MsgBox("Va a introducir un nuevo " & strTipo & ". ¿Está seguro?", vbQuestion + vbYesNo, App.Title) = vbYes Then
            greactivopr = oTipos_reactivo_pr.Insertar
            If greactivopr = 0 Then
                Exit Sub
            End If
        Else
            Exit Sub
        End If
      Else
            If MsgBox("Va a modificar el " & strTipo & ". ¿Está seguro?", vbQuestion + vbYesNo, App.Title) = vbYes Then
            If oTipos_reactivo_pr.Modificar(greactivopr) = False Then
                Exit Sub
            End If
        Else
            Exit Sub
        End If
      End If
        ' ------------------
        ' Componentes
        Dim oReactivos_Componentes_pr As New clsRPR_Componentes
        oReactivos_Componentes_pr.Eliminar (greactivopr) ' Se eliminan todos los componentes del reactivo
        Dim lngFila As Integer, lngTotalFilas As Long
        'lngTotalFilas = total_filas_array - 1
        lngTotalFilas = total_filas_array
        If lngTotalFilas > 0 Then ' Si hay componentes
            'For lngFila = 0 To lngTotalFilas
            For lngFila = 0 To lngTotalFilas - 1
                With oReactivos_Componentes_pr
                    .setTIPO_REACTIVO_PR_ID = greactivopr
                    .setTIPO_REACTIVO_EX_ID = xP(lngFila, ColsP.ID)
                    .setORDEN = lngFila
                    .setPROCEDIMIENTO_REFERENCIA = xP(lngFila, ColsP.P_REFERENCIA)
                    .setCANTIDAD = xP(lngFila, ColsP.cantidad)
                    If IsEmpty(xP(lngFila, ColsP.id_unidad)) Then
                        .setUNIDAD_ID = 0
                    ElseIf Trim(xP(lngFila, ColsP.id_unidad)) = "" Then
                        .setUNIDAD_ID = 0
                    Else
                        .setUNIDAD_ID = xP(lngFila, ColsP.id_unidad)
                    End If
                    .setTIPO = xP(lngFila, ColsP.tipo)
                    .Insertar
                End With
            Next
        End If
        ' -------------
      If auxiliar = 0 Then
            MsgBox "El " & strTipo & " se ha insertado correctamente.", vbOKOnly + vbInformation, App.Title
      Else
            MsgBox "El " & strTipo & " se ha modificado correctamente.", vbOKOnly + vbInformation, App.Title
      End If
      Unload Me
    End If

   On Error GoTo 0
   Exit Sub

cmdok_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure cmdok_Click of Formulario frmRPR_Reactivo"
End Sub
'E0150-F

Private Sub cmdReset_Click()
    If total_filas_array > 0 Then
        Call eliminar_reactivo
    End If
End Sub

Private Sub gridComponentes_DblClick()
   On Error GoTo gridComponentes_DblClick_Error

    If total_filas_array > 0 Then
        If xP(gridComponentes.Bookmark, ColsP.tipo) = "E" Then
            frmREX_Reactivo.PK = xP(gridComponentes.Bookmark, ColsP.ID)
            frmREX_Reactivo.Show 1
        End If
    End If

   On Error GoTo 0
   Exit Sub

gridComponentes_DblClick_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure gridComponentes_DblClick of Formulario frmRPR_Reactivo"
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = 27 Then
        cmdcancel_Click
    End If
End Sub

Private Sub Form_Load()
    log (Me.Name)
    cargar_botones Me
    inicializar_grid
    llenar_combo cmbReactivos, New clsTipos_reactivo_ex, 0, Me, " ANULADO = 0 "
    llenar_combo cmbPropio, New clsRPR_Tipos, 0, Me, ""
    cargar_combo cmbcad, New clsTipos_caducidad
    cargar_combo_unidades
    cargar_combo cmbCentro, New clsCentros
    If greactivopr <> 0 Then
        lbltitulo(0) = "Modificación de Reactivo Propio/Suministro"
'        Label1(2).BackColor = &H80C0FF
        cargar_ReactivoPr
    End If
End Sub


Private Sub tUnidades_DropDownClose()
    gridComponentes.Columns(ColsP.Unidad) = tUnidades.Columns(0)
    gridComponentes.Columns(ColsP.id_unidad) = tUnidades.Columns(1)
    gridComponentes.Col = 0
    gridComponentes.Row = gridComponentes.Row + 1
End Sub

Private Sub txtdatos_GotFocus(Index As Integer)
    txtDatos(Index).BackColor = &H80C0FF
End Sub
Private Sub txtdatos_LostFocus(Index As Integer)
    txtDatos(Index).BackColor = vbWhite
End Sub
Public Sub cargar_ReactivoPr()
    Dim oTipos_reactivo_pr As New clsRPR_Tipos
    
    With oTipos_reactivo_pr
        If .CARGAR(CLng(greactivopr)) = True Then
            txtDatos(0) = .getCODIGO
            txtDatos(1) = .getNOMBRE
            txtDatos(2) = .getALMACENAMIENTO
            txtDatos(3) = .getEQUIPOS
            txtDatos(4) = .getPROTOCOLO
            txtDatos(5) = .getPROC_REFERENCIA
            txtDatos(6) = .getCANTIDAD
            cmbcad.BoundText = .getTIPO_CADUCIDAD_ID
            optTipo(.getTIPO).value = True
            cmbCentro.BoundText = .getCENTRO_ID
         ' Componentes
         Dim oReactivos_Componentes_pr As New clsRPR_Componentes
         Dim rs As ADODB.Recordset
         Set rs = oReactivos_Componentes_pr.Componentes(greactivopr)
         If rs.RecordCount <> 0 Then
            Do
                    Call anadir_reactivo(rs(0), rs(1), rs(2), rs(3), rs(4), rs(5), rs(6))
                rs.MoveNext
            Loop Until rs.EOF
         End If
     End If
    End With
    Set oTipos_reactivo_pr = Nothing
End Sub
Public Function datos_correctos() As Boolean
    datos_correctos = True
    If Trim(txtDatos(0)) = "" Then
        MsgBox "Debe darle un código al Reactivo.", vbInformation, App.Title
        datos_correctos = False
        Exit Function
    End If
    If Trim(txtDatos(1)) = "" Then
        MsgBox "Debe darle un nombre al Reactivo.", vbInformation, App.Title
        datos_correctos = False
        Exit Function
    End If
    If cmbcad.Text = "" Then
        MsgBox "Debe darle una caducidad al Reactivo.", vbInformation, App.Title
        datos_correctos = False
        Exit Function
    End If
    If cmbCentro.Text = "" Then
        MsgBox "Debe indicar un CENTRO.", vbInformation, App.Title
        cmbCentro.SetFocus
        datos_correctos = False
        Exit Function
    End If
End Function
Private Sub inicializar_grid()
    On Error GoTo inicializar_grid_Error
    
    gridComponentes.Col = 0
    gridComponentes.Row = 0
    xP.Clear
    xP.ReDim 0, filasP, 0, ColP
    xP.Clear
    Set gridComponentes.Array = xP
    gridComponentes.Refresh
    
    On Error GoTo 0
    
    Exit Sub
    
inicializar_grid_Error:
    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure inicializar_grid of Formulario frmRPR_Reactivo"
End Sub

' Procedimiento que añade un reactivo al grid
Private Sub anadir_reactivo(strReactivo As String, strId As String, PROCEDIMIENTO As String, cantidad As String, Unidad As String, id_unidad As String, tipo As String)
    Dim lngNuevaFila As Long
    
    lngNuevaFila = total_filas_array()
    xP(lngNuevaFila, ColsP.Reactivo) = strReactivo
    xP(lngNuevaFila, ColsP.P_REFERENCIA) = PROCEDIMIENTO
    xP(lngNuevaFila, ColsP.cantidad) = cantidad
    xP(lngNuevaFila, ColsP.Unidad) = Unidad
    xP(lngNuevaFila, ColsP.id_unidad) = id_unidad
    xP(lngNuevaFila, ColsP.ID) = strId
    xP(lngNuevaFila, ColsP.tipo) = tipo
End Sub

' Procedimiento que elimina el reactivo seleccionado del grid
Private Sub eliminar_reactivo()
    On Error Resume Next
    Dim lngCol As Long, lngFila As Long
    
    ' Se borra la fila seleccionada
    For lngCol = 0 To ColP
        gridComponentes.SelBookmarks.Add gridComponentes.Bookmark
        xP(gridComponentes.Bookmark, lngCol) = ""
        gridComponentes.SelBookmarks.Remove 0
    Next lngCol
    ' Mover hacia arriba los eltos que queden por debajo del borrado
    For lngFila = gridComponentes.Bookmark To xP.UpperBound(1)
        If xP(lngFila, 0) <> "" Then
            Call copiar_fila(lngFila, lngFila - 1)
            Call borrar_fila(lngFila)
        End If
    Next lngFila
    Set gridComponentes.Array = xP
    gridComponentes.Refresh
End Sub
' Procedimiento que mueve los datos del tdbgrid de la fila orig a la fila dest
Private Sub copiar_fila(lngFilaOrig As Long, lngFilaDest As Long)
    xP(lngFilaDest, ColsP.Reactivo) = xP(lngFilaOrig, ColsP.Reactivo)
    xP(lngFilaDest, ColsP.P_REFERENCIA) = xP(lngFilaOrig, ColsP.P_REFERENCIA)
    xP(lngFilaDest, ColsP.ID) = xP(lngFilaOrig, ColsP.ID)
End Sub

' Procedimiento que borra los datos del tdbgrid de la fila
Private Sub borrar_fila(lngFila As Long)
    xP(lngFila, ColsP.Reactivo) = ""
    xP(lngFila, ColsP.P_REFERENCIA) = ""
    xP(lngFila, ColsP.Unidad) = ""
    xP(lngFila, ColsP.ID) = ""
End Sub

' Función que devuelve el número de filas (rellenas) que hay en el array
Private Function total_filas_array() As Long
    Dim lngFila As Long
    
    lngFila = 0
    While Not xP(lngFila, 0) = ""
        lngFila = lngFila + 1
    Wend
    total_filas_array = lngFila
End Function
Private Sub cargar_combo_unidades()
    Dim ounidades As New clsUnidades
    Dim rs As ADODB.Recordset
    Set rs = ounidades.Listado("")
    Dim i As Integer
    If rs.RecordCount > 0 Then
        xUnidades.ReDim 0, rs.RecordCount, 0, ColUnidades
        xUnidades.Clear
        i = 0
        Do
            xUnidades(i, 0) = CStr(rs("NOMBRE"))
            xUnidades(i, 1) = CStr(rs("ID_UNIDAD"))
            i = i + 1
            rs.MoveNext
        Loop Until rs.EOF
    Else
        xUnidades.ReDim 0, 0, 0, ColUnidades
    End If
    Set tUnidades.Array = xUnidades
    tUnidades.Refresh
End Sub

