VERSION 5.00
Object = "{77DDF307-D82B-4757-8B3A-106EC9D75D4B}#8.0#0"; "tdbg8.ocx"
Begin VB.Form frmPlasma_Traccion 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "TENSILE STRENGTH"
   ClientHeight    =   8100
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   12135
   Icon            =   "frmRecepcion.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8100
   ScaleWidth      =   12135
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame6 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   "ADHESIVE THICKNESS:"
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
      Left            =   4410
      TabIndex        =   29
      Top             =   6255
      Width           =   7665
      Begin VB.TextBox txtEspesor 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   135
         MaxLength       =   255
         TabIndex        =   14
         Top             =   225
         Width           =   7440
      End
   End
   Begin VB.CheckBox chkModificar 
      Caption         =   "Permiso Modificar Cerrada"
      Height          =   195
      Left            =   5445
      TabIndex        =   28
      Top             =   7695
      Visible         =   0   'False
      Width           =   1230
   End
   Begin VB.Frame Frame5 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   "ADHESIVE APPLIED:"
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
      Left            =   4410
      TabIndex        =   27
      Top             =   5490
      Width           =   7665
      Begin VB.TextBox txtAdhesive 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   135
         MaxLength       =   255
         TabIndex        =   13
         Top             =   225
         Width           =   7440
      End
   End
   Begin VB.Frame Frame4 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   "EVALUATION:"
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
      Height          =   2130
      Left            =   45
      TabIndex        =   18
      Top             =   3240
      Width           =   4290
      Begin VB.TextBox txtT4 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2025
         Locked          =   -1  'True
         MaxLength       =   255
         TabIndex        =   25
         Top             =   1575
         Width           =   1995
      End
      Begin VB.TextBox txtT3 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2025
         Locked          =   -1  'True
         MaxLength       =   255
         TabIndex        =   23
         Top             =   1170
         Width           =   1995
      End
      Begin VB.TextBox txtT2 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2025
         Locked          =   -1  'True
         MaxLength       =   255
         TabIndex        =   21
         Top             =   765
         Width           =   1995
      End
      Begin VB.TextBox txtT1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2025
         Locked          =   -1  'True
         MaxLength       =   255
         TabIndex        =   19
         Top             =   360
         Width           =   1995
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "[(T1+T2)/2]-2.66(T1-T2) :"
         Height          =   195
         Index           =   3
         Left            =   135
         TabIndex        =   26
         Top             =   1665
         Width           =   1800
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "T3 (Minimum Value) : "
         Height          =   195
         Index           =   2
         Left            =   135
         TabIndex        =   24
         Top             =   1260
         Width           =   1530
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "T2 (Median) : "
         Height          =   195
         Index           =   1
         Left            =   135
         TabIndex        =   22
         Top             =   855
         Width           =   990
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "T1 (Maximum Value) : "
         Height          =   195
         Index           =   0
         Left            =   135
         TabIndex        =   20
         Top             =   450
         Width           =   1575
      End
   End
   Begin VB.Frame Frame3 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   "TEST VELOCITY:"
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
      Height          =   735
      Left            =   4410
      TabIndex        =   17
      Top             =   4005
      Width           =   7665
      Begin VB.TextBox txtVelocity 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   135
         MaxLength       =   255
         TabIndex        =   11
         Top             =   225
         Width           =   7440
      End
   End
   Begin VB.Frame Frame2 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   "EPOXY PASTE ADHESIVE BATCH:"
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
      Left            =   4410
      TabIndex        =   15
      Top             =   4770
      Width           =   7665
      Begin VB.TextBox txtEpoxy 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   135
         MaxLength       =   255
         TabIndex        =   12
         Top             =   225
         Width           =   7440
      End
   End
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   "ROOM CONDITIONS:"
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
      Height          =   735
      Left            =   4410
      TabIndex        =   9
      Top             =   3240
      Width           =   7665
      Begin VB.TextBox txtRoom 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   135
         MaxLength       =   255
         TabIndex        =   10
         Top             =   225
         Width           =   7440
      End
   End
   Begin VB.Frame frmBondTraccion 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   "RESULTS:"
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
      Height          =   1500
      Left            =   45
      TabIndex        =   4
      Top             =   5445
      Width           =   4290
      Begin VB.TextBox txtMedia 
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
         Height          =   375
         Left            =   2025
         Locked          =   -1  'True
         MaxLength       =   255
         TabIndex        =   6
         Top             =   360
         Width           =   1995
      End
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
         Height          =   375
         Left            =   2025
         Locked          =   -1  'True
         MaxLength       =   255
         TabIndex        =   5
         Top             =   900
         Width           =   1995
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "AVERAGE:"
         Height          =   195
         Index           =   12
         Left            =   135
         TabIndex        =   8
         Top             =   450
         Width           =   810
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "S.D.:"
         Height          =   195
         Index           =   31
         Left            =   135
         TabIndex        =   7
         Top             =   945
         Width           =   360
      End
   End
   Begin VB.CommandButton cmdok 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Aceptar"
      Height          =   840
      Left            =   9765
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   7155
      Width           =   1140
   End
   Begin VB.CommandButton cmdSalir 
      BackColor       =   &H00E0E0E0&
      Caption         =   "ESC-Salir"
      Height          =   840
      Left            =   10935
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   7155
      Width           =   1140
   End
   Begin TrueDBGrid80.TDBGrid grid 
      Height          =   2790
      Left            =   0
      TabIndex        =   1
      Top             =   360
      Width           =   12075
      _ExtentX        =   21299
      _ExtentY        =   4921
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
      Columns(1).Caption=   "Diameter (inch)"
      Columns(1).DataField=   ""
      Columns(1).NumberFormat=   "0.000"
      Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(2)._VlistStyle=   0
      Columns(2)._MaxComboItems=   5
      Columns(2).Caption=   "Area (inch ^ 2)"
      Columns(2).DataField=   ""
      Columns(2).NumberFormat=   "0.000"
      Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(3)._VlistStyle=   0
      Columns(3)._MaxComboItems=   5
      Columns(3).Caption=   "Load (p)"
      Columns(3).DataField=   ""
      Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(4)._VlistStyle=   0
      Columns(4)._MaxComboItems=   5
      Columns(4).Caption=   "Tensile Strength (psi)"
      Columns(4).DataField=   ""
      Columns(4)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(5)._VlistStyle=   0
      Columns(5)._MaxComboItems=   5
      Columns(5).Caption=   "Failure location"
      Columns(5).DataField=   ""
      Columns(5)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns.Count   =   6
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
      Splits(0)._ColumnProps(0)=   "Columns.Count=6"
      Splits(0)._ColumnProps(1)=   "Column(0).Width=2884"
      Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
      Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=2805"
      Splits(0)._ColumnProps(4)=   "Column(0)._EditAlways=0"
      Splits(0)._ColumnProps(5)=   "Column(0)._ColStyle=8193"
      Splits(0)._ColumnProps(6)=   "Column(0).AllowFocus=0"
      Splits(0)._ColumnProps(7)=   "Column(0).Order=1"
      Splits(0)._ColumnProps(8)=   "Column(1).Width=2408"
      Splits(0)._ColumnProps(9)=   "Column(1).DividerColor=0"
      Splits(0)._ColumnProps(10)=   "Column(1)._WidthInPix=2328"
      Splits(0)._ColumnProps(11)=   "Column(1)._EditAlways=0"
      Splits(0)._ColumnProps(12)=   "Column(1)._ColStyle=1"
      Splits(0)._ColumnProps(13)=   "Column(1).Order=2"
      Splits(0)._ColumnProps(14)=   "Column(2).Width=2434"
      Splits(0)._ColumnProps(15)=   "Column(2).DividerColor=0"
      Splits(0)._ColumnProps(16)=   "Column(2)._WidthInPix=2355"
      Splits(0)._ColumnProps(17)=   "Column(2)._EditAlways=0"
      Splits(0)._ColumnProps(18)=   "Column(2)._ColStyle=8193"
      Splits(0)._ColumnProps(19)=   "Column(2).AllowFocus=0"
      Splits(0)._ColumnProps(20)=   "Column(2).Order=3"
      Splits(0)._ColumnProps(21)=   "Column(3).Width=2328"
      Splits(0)._ColumnProps(22)=   "Column(3).DividerColor=0"
      Splits(0)._ColumnProps(23)=   "Column(3)._WidthInPix=2249"
      Splits(0)._ColumnProps(24)=   "Column(3)._EditAlways=0"
      Splits(0)._ColumnProps(25)=   "Column(3)._ColStyle=1"
      Splits(0)._ColumnProps(26)=   "Column(3).Order=4"
      Splits(0)._ColumnProps(27)=   "Column(4).Width=2566"
      Splits(0)._ColumnProps(28)=   "Column(4).DividerColor=0"
      Splits(0)._ColumnProps(29)=   "Column(4)._WidthInPix=2487"
      Splits(0)._ColumnProps(30)=   "Column(4)._EditAlways=0"
      Splits(0)._ColumnProps(31)=   "Column(4)._ColStyle=8193"
      Splits(0)._ColumnProps(32)=   "Column(4).AllowFocus=0"
      Splits(0)._ColumnProps(33)=   "Column(4).Order=5"
      Splits(0)._ColumnProps(34)=   "Column(5).Width=2725"
      Splits(0)._ColumnProps(35)=   "Column(5).DividerColor=0"
      Splits(0)._ColumnProps(36)=   "Column(5)._WidthInPix=2646"
      Splits(0)._ColumnProps(37)=   "Column(5)._EditAlways=0"
      Splits(0)._ColumnProps(38)=   "Column(5).Order=6"
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
      _StyleDefs(36)  =   "Splits(0).Columns(0).Style:id=24,.parent=11,.alignment=2,.bgcolor=&HD7D7D7&"
      _StyleDefs(37)  =   ":id=24,.locked=-1,.bold=-1,.fontsize=825,.italic=0,.underline=0,.strikethrough=0"
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
      _StyleDefs(48)  =   "Splits(0).Columns(2).Style:id=54,.parent=11,.alignment=2,.bgcolor=&HD7D7D7&"
      _StyleDefs(49)  =   ":id=54,.locked=-1"
      _StyleDefs(50)  =   "Splits(0).Columns(2).HeadingStyle:id=51,.parent=12"
      _StyleDefs(51)  =   "Splits(0).Columns(2).FooterStyle:id=52,.parent=13"
      _StyleDefs(52)  =   "Splits(0).Columns(2).EditorStyle:id=53,.parent=15"
      _StyleDefs(53)  =   "Splits(0).Columns(3).Style:id=32,.parent=11,.alignment=2"
      _StyleDefs(54)  =   "Splits(0).Columns(3).HeadingStyle:id=29,.parent=12"
      _StyleDefs(55)  =   "Splits(0).Columns(3).FooterStyle:id=30,.parent=13"
      _StyleDefs(56)  =   "Splits(0).Columns(3).EditorStyle:id=31,.parent=15"
      _StyleDefs(57)  =   "Splits(0).Columns(4).Style:id=36,.parent=11,.alignment=2,.bgcolor=&HD7D7D7&"
      _StyleDefs(58)  =   ":id=36,.locked=-1"
      _StyleDefs(59)  =   "Splits(0).Columns(4).HeadingStyle:id=33,.parent=12"
      _StyleDefs(60)  =   "Splits(0).Columns(4).FooterStyle:id=34,.parent=13"
      _StyleDefs(61)  =   "Splits(0).Columns(4).EditorStyle:id=35,.parent=15"
      _StyleDefs(62)  =   "Splits(0).Columns(5).Style:id=58,.parent=11"
      _StyleDefs(63)  =   "Splits(0).Columns(5).HeadingStyle:id=55,.parent=12"
      _StyleDefs(64)  =   "Splits(0).Columns(5).FooterStyle:id=56,.parent=13"
      _StyleDefs(65)  =   "Splits(0).Columns(5).EditorStyle:id=57,.parent=15"
      _StyleDefs(66)  =   "Named:id=37:Normal"
      _StyleDefs(67)  =   ":id=37,.parent=0,.alignment=3,.bgcolor=&H80000018&,.bgpicMode=2,.bold=0"
      _StyleDefs(68)  =   ":id=37,.fontsize=825,.italic=0,.underline=0,.strikethrough=0,.charset=0"
      _StyleDefs(69)  =   ":id=37,.fontname=MS Sans Serif"
      _StyleDefs(70)  =   "Named:id=38:Heading"
      _StyleDefs(71)  =   ":id=38,.parent=37,.valignment=2,.bgcolor=&H80000016&,.fgcolor=&H80000012&"
      _StyleDefs(72)  =   ":id=38,.wraptext=-1,.bold=0,.fontsize=825,.italic=0,.underline=0"
      _StyleDefs(73)  =   ":id=38,.strikethrough=0,.charset=0"
      _StyleDefs(74)  =   ":id=38,.fontname=MS Sans Serif"
      _StyleDefs(75)  =   "Named:id=39:Footing"
      _StyleDefs(76)  =   ":id=39,.parent=37,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(77)  =   "Named:id=40:Selected"
      _StyleDefs(78)  =   ":id=40,.parent=37,.bgcolor=&H8080FF&,.fgcolor=&H0&,.bold=0,.fontsize=825"
      _StyleDefs(79)  =   ":id=40,.italic=0,.underline=0,.strikethrough=0,.charset=0"
      _StyleDefs(80)  =   ":id=40,.fontname=MS Sans Serif"
      _StyleDefs(81)  =   "Named:id=41:Caption"
      _StyleDefs(82)  =   ":id=41,.parent=38,.alignment=2"
      _StyleDefs(83)  =   "Named:id=42:HighlightRow"
      _StyleDefs(84)  =   ":id=42,.parent=37,.bgcolor=&H80000008&,.fgcolor=&H80000005&"
      _StyleDefs(85)  =   "Named:id=43:EvenRow"
      _StyleDefs(86)  =   ":id=43,.parent=37,.bgcolor=&HFFFFFF&,.wraptext=-1"
      _StyleDefs(87)  =   "Named:id=44:OddRow"
      _StyleDefs(88)  =   ":id=44,.parent=37,.bgcolor=&HD5ECF9&"
      _StyleDefs(89)  =   "Named:id=47:RecordSelector"
      _StyleDefs(90)  =   ":id=47,.parent=38"
      _StyleDefs(91)  =   "Named:id=50:FilterBar"
      _StyleDefs(92)  =   ":id=50,.parent=37"
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
      Left            =   10305
      TabIndex        =   16
      Top             =   0
      Width           =   1770
   End
   Begin VB.Label lbltitulo 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Caption         =   "TENSILE STRENGTH"
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
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   12105
   End
End
Attribute VB_Name = "frmPlasma_Traccion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public MUESTRA_ID As Long
Public tipo As Integer

Dim xTRACCION As New XArrayDB
Const filasGrid As Integer = 4
Const ColGrid As Integer = 6
Private Enum ColsGrid
    IDENTIFICATION = 0
    DIAMETER = 1
    AREA = 2
    LOADP = 3
    TENSILE = 4
    LOCATION = 5
End Enum

Private Sub cmdok_Click()
    Dim i As Integer
    Dim oPTP As New clsPlasma_traccion_p
   On Error GoTo cmdok_Click_Error

    For i = 0 To filasGrid
        If Not IsEmpty(xTRACCION(i, ColsGrid.IDENTIFICATION)) Then
          If Trim(xTRACCION(i, ColsGrid.TENSILE)) <> "" Then
            With oPTP
                .setMUESTRA_ID = MUESTRA_ID
                .setTIPO = tipo
                .setORDEN = i
                If Trim(xTRACCION(i, ColsGrid.IDENTIFICATION)) <> "" Then
                    .setIDENTIFICATION = xTRACCION(i, ColsGrid.IDENTIFICATION)
                Else
                    .setIDENTIFICATION = "0"
                End If
                If Trim(xTRACCION(i, ColsGrid.DIAMETER)) <> "" Then
                    .setDIAMETER = Replace(xTRACCION(i, ColsGrid.DIAMETER), ",", ".")
                Else
                    .setDIAMETER = "0"
                End If
                If Trim(xTRACCION(i, ColsGrid.LOADP)) <> "" Then
                    .setAREA = Replace(xTRACCION(i, ColsGrid.AREA), ",", ".")
                Else
                    .setAREA = "0"
                End If
                If Trim(xTRACCION(i, ColsGrid.LOADP)) <> "" Then
                    .setLOADP = Replace(xTRACCION(i, ColsGrid.LOADP), ",", ".")
                Else
                    .setLOADP = "0"
                End If
                .setTENSILE = Replace(xTRACCION(i, ColsGrid.TENSILE), ",", ".")
                .setLOCATION = xTRACCION(i, ColsGrid.LOCATION)
                
                .Insertar
            End With
          End If
        End If
    Next
    ' Pasar datos
    Dim res As String
'    For i = 0 To 2
'        If Not IsEmpty(xTRACCION(i, ColsGrid.IDENTIFICATION)) Then
'            If Trim(xTRACCION(i, ColsGrid.TENSILE)) <> "" Then
'                If res <> "" Then
'                    res = res & "-"
'                End If
'                res = res & Trim(xTRACCION(i, ColsGrid.TENSILE))
'            End If
'        End If
'    Next
    res = "Max: " & txtT1 & " psi. Min: " & txtT3 & " psi."
    If tipo = 1 Then
        frmPlasma_Resultados.txtDatos(32) = res
        frmPlasma_Resultados.txtDatos(33) = txtMedia
    ElseIf tipo = 2 Then
        frmPlasma_Resultados.txtDatos(42) = res
        frmPlasma_Resultados.txtDatos(43) = txtMedia
    End If
    Dim opT As New clsPlasma_traccion
    With opT
        .setMUESTRA_ID = MUESTRA_ID
        .setTIPO = tipo
        .setROOM = txtRoom
        .setVELOCITY = txtVelocity
        .setEPOXY = txtEpoxy
        .setADHESIVE = txtAdhesive
        .setESPESOR = txtEspesor
        .setAVERAGE = txtMedia
        .setSD = txtSD
        .Insertar
    End With
    Dim oPR As New clsPlasma_resultados
    oPR.actualizarTraccion MUESTRA_ID, tipo, res, txtMedia
    Set oPR = Nothing
    
    Unload Me

   On Error GoTo 0
   Exit Sub

cmdok_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure cmdok_Click of Formulario frmPlasma_Traccion"
End Sub

Private Sub cmdSalir_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    log (Me.Name)
    cargar_botones Me
    inicializar_grid
    permisos
    activarCampos
    calcularResultados
End Sub
Private Sub activarCampos()
    Dim oMuestra As New clsMuestra
    oMuestra.CargaMuestra MUESTRA_ID
    Select Case oMuestra.getCERRADA
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
    If oMuestra.getCERRADA <> 0 And chkModificar.Value = Unchecked Then
        cmdok.Enabled = False
        grid.Enabled = False
        txtRoom.Enabled = False
        txtVelocity.Enabled = False
        txtEpoxy.Enabled = False
        txtAdhesive.Enabled = False
        txtEspesor.Enabled = False
    End If
End Sub
Private Sub inicializar_grid()
   On Error GoTo inicializar_grid_Error

    xTRACCION.Clear
    xTRACCION.ReDim 0, filasGrid, 0, ColGrid
    xTRACCION.Clear
    ' Configuración inicial del grid
    xTRACCION(0, ColsGrid.IDENTIFICATION) = "P1"
    xTRACCION(1, ColsGrid.IDENTIFICATION) = "P2"
    xTRACCION(2, ColsGrid.IDENTIFICATION) = "P3"
    xTRACCION(3, ColsGrid.IDENTIFICATION) = "B (Blank Strength)"
    
    xTRACCION(0, ColsGrid.LOCATION) = "Intracoating: -- % / In epoxy : -- %"
    xTRACCION(1, ColsGrid.LOCATION) = "Intracoating: -- % / In epoxy : -- %"
    xTRACCION(2, ColsGrid.LOCATION) = "Intracoating: -- % / In epoxy : -- %"
    
    txtRoom = "-- ºC / -- %Hr"
    ' Cargar de los datos de la muestra
    Dim rs As ADODB.Recordset
    Dim oPTP As New clsPlasma_traccion_p
    Set rs = oPTP.Listado(MUESTRA_ID, tipo)
    If rs.RecordCount > 0 Then
        Do
            xTRACCION(rs("ORDEN"), ColsGrid.DIAMETER) = CStr(rs("DIAMETER"))
            xTRACCION(rs("ORDEN"), ColsGrid.AREA) = CStr(rs("AREA"))
            xTRACCION(rs("ORDEN"), ColsGrid.LOADP) = CStr(rs("LOADP"))
            xTRACCION(rs("ORDEN"), ColsGrid.TENSILE) = CStr(rs("TENSILE"))
            xTRACCION(rs("ORDEN"), ColsGrid.LOCATION) = CStr(rs("LOCATION"))
            
            rs.MoveNext
        Loop Until rs.EOF
    End If
    Set rs = Nothing
    Set grid.Array = xTRACCION
    grid.Refresh
    Dim opT As New clsPlasma_traccion
    If opT.Carga(MUESTRA_ID, tipo) = True Then
        txtMedia = opT.getAVERAGE
        txtSD = opT.getSD
        txtRoom = opT.getROOM
        txtVelocity = opT.getVELOCITY
        txtEpoxy = opT.getEPOXY
        txtAdhesive = opT.getADHESIVE
        txtEspesor = opT.getESPESOR
    End If
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
    Dim total As Single
    Dim CANTIDAD As Integer
    Dim sumatorio As Single
    Dim medida As Single
    Dim numero_medidas As Integer
    Dim RESULTADO As Single
   On Error GoTo calcularDesviacion_Error
    txtMedia = ""
    txtSD = ""
    media = 0
    sumatorio = 0
    numero_medidas = 0
    Dim i As Integer
    ' Calculo de Area y Tensile
    For i = 0 To 3
        If Not IsEmpty(xTRACCION(i, ColsGrid.IDENTIFICATION)) Then
            If Trim(xTRACCION(i, ColsGrid.DIAMETER)) <> "" Then
                xTRACCION(i, ColsGrid.AREA) = PI * (CSng(xTRACCION(i, ColsGrid.DIAMETER) ^ 2) / 4)
            End If
            If Trim(xTRACCION(i, ColsGrid.LOADP)) <> "" And Trim(xTRACCION(i, ColsGrid.AREA)) <> "" Then
                xTRACCION(i, ColsGrid.TENSILE) = CInt(CSng(xTRACCION(i, ColsGrid.LOADP)) / CSng(xTRACCION(i, ColsGrid.AREA)))
            End If
        End If
    Next
    grid.Refresh
    ' Montamos los resultados separados por -
    Dim res As String
    For i = 0 To 2
        If Not IsEmpty(xTRACCION(i, ColsGrid.IDENTIFICATION)) Then
            If Trim(xTRACCION(i, ColsGrid.TENSILE)) <> "" Then
                If res <> "" Then
                    res = res & "-"
                End If
                res = res & Trim(xTRACCION(i, ColsGrid.TENSILE))
            End If
        End If
    Next
    If res = "" Then
        Exit Sub
    End If
    lista = Split(res, "-")
    ' MEDIA
    For i = LBound(lista) To UBound(lista)
        If IsNumeric(lista(i)) Then
            total = total + lista(i)
            CANTIDAD = CANTIDAD + 1
        End If
    Next
    If CANTIDAD > 0 Then
        media = CInt(total / CANTIDAD)
    End If
'    txtMedia = CStr(media) & " psi"
    ' DESVIACION
    If UBound(lista) < 2 Then
        Exit Sub
    End If
    For i = LBound(lista) To UBound(lista)
        If IsNumeric(lista(i)) Then
            medida = lista(i)
            sumatorio = sumatorio + ((medida - media) * (medida - media))
            numero_medidas = numero_medidas + 1
        End If
    Next
    txtSD = formatear(Sqr(sumatorio / (numero_medidas - 1)), 5, 1)

    ' NUEVA MEDIA
    Dim arr(3) As Integer
    arr(0) = xTRACCION(0, ColsGrid.TENSILE)
    arr(1) = xTRACCION(1, ColsGrid.TENSILE)
    arr(2) = xTRACCION(2, ColsGrid.TENSILE)
    
    For i = 0 To 2 Step 1
        For j = 0 To (2 - 1) Step 1
            If arr(j) > arr(j + 1) Then ' Para Descendente, Inviertes el > con <
                aux = arr(j + 1)
                arr(j + 1) = arr(j)
                arr(j) = aux
            End If
        Next j
    Next i
    txtT1 = arr(2)
    txtT2 = Mediana(arr)
    txtT3 = arr(0)
    txtT4 = CInt(((CInt(txtT1) + CInt(txtT2)) / 2) - 2.66 * (CInt(txtT1) - CInt(txtT2)))
    
    If CInt(txtT1) < CInt(txtT2) Then
        txtMedia = CStr(CInt((txtT1 + txtT2) / 2)) & " psi"
    Else
        txtMedia = CStr(media) & " psi"
    End If
   On Error GoTo 0
   Exit Sub

calcularDesviacion_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure calcularDesviacion of Formulario frmPlasma_Resultados"
End Sub
Private Function Mediana(ByRef arr() As Integer) As Integer
    Dim lngElementos As Long, lngMedio As Long
    lngElementos = UBound(arr) - LBound(arr)
    lngMedio = LBound(arr) + (lngElementos \ 2)
    If lngElementos And 1 Then
        Mediana = arr(lngMedio)
    Else
        Mediana = (arr(lngMedio) + arr(lngMedio - 1)) / 2
    End If
End Function

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
