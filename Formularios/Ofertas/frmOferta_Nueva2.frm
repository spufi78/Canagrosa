VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{CEA3B7F6-2847-4E5E-A551-DB7A62489D44}#46.0#0"; "miCombo.ocx"
Object = "{77DDF307-D82B-4757-8B3A-106EC9D75D4B}#8.0#0"; "tdbg8.ocx"
Begin VB.Form frmOferta_Nueva2 
   BackColor       =   &H00C0C0C0&
   Caption         =   "Gestión de Ofertas"
   ClientHeight    =   9510
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   13935
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmOferta_Nueva2.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   9510
   ScaleWidth      =   13935
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdRequisitos 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Requisitos"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   915
      Left            =   6750
      Picture         =   "frmOferta_Nueva2.frx":08CA
      Style           =   1  'Graphical
      TabIndex        =   58
      Top             =   8550
      Visible         =   0   'False
      Width           =   1995
   End
   Begin VB.CommandButton cmdHistorialCambios 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Seguimiento"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   915
      Left            =   4725
      Style           =   1  'Graphical
      TabIndex        =   50
      Top             =   8550
      Visible         =   0   'False
      Width           =   1995
   End
   Begin TrueDBGrid80.TDBGrid grid 
      Height          =   3735
      Left            =   45
      TabIndex        =   14
      Top             =   4365
      Width           =   13785
      _ExtentX        =   24315
      _ExtentY        =   6588
      _LayoutType     =   4
      _RowHeight      =   -2147483647
      _WasPersistedAsPixels=   0
      Columns(0)._VlistStyle=   0
      Columns(0)._MaxComboItems=   5
      Columns(0).Caption=   "NºGeneral"
      Columns(0).DataField=   ""
      Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(1)._VlistStyle=   0
      Columns(1)._MaxComboItems=   5
      Columns(1).Caption=   "Fecha"
      Columns(1).DataField=   ""
      Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(2)._VlistStyle=   0
      Columns(2)._MaxComboItems=   5
      Columns(2).Caption=   "NºParticular"
      Columns(2).DataField=   ""
      Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(3)._VlistStyle=   0
      Columns(3)._MaxComboItems=   5
      Columns(3).Caption=   "Referencia Cliente"
      Columns(3).DataField=   ""
      Columns(3).NumberFormat=   "Currency"
      Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns.Count   =   4
      Splits(0)._UserFlags=   0
      Splits(0).ExtendRightColumn=   -1  'True
      Splits(0).ShowCollapseExpandIcons=   0   'False
      Splits(0).MarqueeStyle=   1
      Splits(0).AllowRowSizing=   0   'False
      Splits(0).RecordSelectors=   0   'False
      Splits(0).RecordSelectorWidth=   503
      Splits(0).AllowColSelect=   0   'False
      Splits(0).DividerColor=   12632256
      Splits(0).SpringMode=   0   'False
      Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
      Splits(0)._ColumnProps(0)=   "Columns.Count=4"
      Splits(0)._ColumnProps(1)=   "Column(0).Width=1826"
      Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
      Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=1746"
      Splits(0)._ColumnProps(4)=   "Column(0)._EditAlways=0"
      Splits(0)._ColumnProps(5)=   "Column(0)._ColStyle=1"
      Splits(0)._ColumnProps(6)=   "Column(0).Order=1"
      Splits(0)._ColumnProps(7)=   "Column(1).Width=1905"
      Splits(0)._ColumnProps(8)=   "Column(1).DividerColor=0"
      Splits(0)._ColumnProps(9)=   "Column(1)._WidthInPix=1826"
      Splits(0)._ColumnProps(10)=   "Column(1)._EditAlways=0"
      Splits(0)._ColumnProps(11)=   "Column(1)._ColStyle=1"
      Splits(0)._ColumnProps(12)=   "Column(1).Order=2"
      Splits(0)._ColumnProps(13)=   "Column(2).Width=2223"
      Splits(0)._ColumnProps(14)=   "Column(2).DividerColor=0"
      Splits(0)._ColumnProps(15)=   "Column(2)._WidthInPix=2143"
      Splits(0)._ColumnProps(16)=   "Column(2)._EditAlways=0"
      Splits(0)._ColumnProps(17)=   "Column(2)._ColStyle=1"
      Splits(0)._ColumnProps(18)=   "Column(2).Order=3"
      Splits(0)._ColumnProps(19)=   "Column(3).Width=6509"
      Splits(0)._ColumnProps(20)=   "Column(3).DividerColor=0"
      Splits(0)._ColumnProps(21)=   "Column(3)._WidthInPix=6429"
      Splits(0)._ColumnProps(22)=   "Column(3)._EditAlways=0"
      Splits(0)._ColumnProps(23)=   "Column(3)._ColStyle=0"
      Splits(0)._ColumnProps(24)=   "Column(3).Order=4"
      Splits.Count    =   1
      PrintInfos(0)._StateFlags=   0
      PrintInfos(0).Name=   "piInternal 0"
      PrintInfos(0).PageHeaderFont=   "Size=8.25,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=MS Sans Serif"
      PrintInfos(0).PageFooterFont=   "Size=8.25,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=MS Sans Serif"
      PrintInfos(0).PageHeaderHeight=   0
      PrintInfos(0).PageFooterHeight=   0
      PrintInfos.Count=   1
      Appearance      =   0
      DataMode        =   4
      DefColWidth     =   0
      EditDropDown    =   0   'False
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
      _StyleDefs(24)  =   "Splits(0).Style:id=67,.parent=1"
      _StyleDefs(25)  =   "Splits(0).CaptionStyle:id=76,.parent=4"
      _StyleDefs(26)  =   "Splits(0).HeadingStyle:id=68,.parent=2,.namedParent=38"
      _StyleDefs(27)  =   "Splits(0).FooterStyle:id=69,.parent=3"
      _StyleDefs(28)  =   "Splits(0).InactiveStyle:id=70,.parent=5"
      _StyleDefs(29)  =   "Splits(0).SelectedStyle:id=72,.parent=6,.namedParent=40"
      _StyleDefs(30)  =   "Splits(0).EditorStyle:id=71,.parent=7,.namedParent=40"
      _StyleDefs(31)  =   "Splits(0).HighlightRowStyle:id=73,.parent=8"
      _StyleDefs(32)  =   "Splits(0).EvenRowStyle:id=74,.parent=9,.namedParent=43"
      _StyleDefs(33)  =   "Splits(0).OddRowStyle:id=75,.parent=10,.namedParent=44"
      _StyleDefs(34)  =   "Splits(0).RecordSelectorStyle:id=77,.parent=45"
      _StyleDefs(35)  =   "Splits(0).FilterBarStyle:id=78,.parent=48"
      _StyleDefs(36)  =   "Splits(0).Columns(0).Style:id=86,.parent=67,.alignment=2,.wraptext=0"
      _StyleDefs(37)  =   "Splits(0).Columns(0).HeadingStyle:id=83,.parent=68"
      _StyleDefs(38)  =   "Splits(0).Columns(0).FooterStyle:id=84,.parent=69"
      _StyleDefs(39)  =   "Splits(0).Columns(0).EditorStyle:id=85,.parent=71,.namedParent=40"
      _StyleDefs(40)  =   "Splits(0).Columns(1).Style:id=90,.parent=67,.alignment=2,.wraptext=0,.bold=0"
      _StyleDefs(41)  =   ":id=90,.fontsize=825,.italic=0,.underline=0,.strikethrough=0,.charset=0"
      _StyleDefs(42)  =   ":id=90,.fontname=MS Sans Serif"
      _StyleDefs(43)  =   "Splits(0).Columns(1).HeadingStyle:id=87,.parent=68"
      _StyleDefs(44)  =   "Splits(0).Columns(1).FooterStyle:id=88,.parent=69"
      _StyleDefs(45)  =   "Splits(0).Columns(1).EditorStyle:id=89,.parent=71,.fgcolor=&H0&,.bold=0"
      _StyleDefs(46)  =   ":id=89,.fontsize=825,.italic=0,.underline=0,.strikethrough=0,.charset=0"
      _StyleDefs(47)  =   ":id=89,.fontname=MS Sans Serif"
      _StyleDefs(48)  =   "Splits(0).Columns(2).Style:id=94,.parent=67,.alignment=2,.wraptext=0"
      _StyleDefs(49)  =   "Splits(0).Columns(2).HeadingStyle:id=91,.parent=68"
      _StyleDefs(50)  =   "Splits(0).Columns(2).FooterStyle:id=92,.parent=69"
      _StyleDefs(51)  =   "Splits(0).Columns(2).EditorStyle:id=93,.parent=71"
      _StyleDefs(52)  =   "Splits(0).Columns(3).Style:id=102,.parent=67,.alignment=0"
      _StyleDefs(53)  =   "Splits(0).Columns(3).HeadingStyle:id=99,.parent=68"
      _StyleDefs(54)  =   "Splits(0).Columns(3).FooterStyle:id=100,.parent=69"
      _StyleDefs(55)  =   "Splits(0).Columns(3).EditorStyle:id=101,.parent=71"
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
      _StyleDefs(74)  =   ":id=43,.parent=37,.bgcolor=&HFFFFFF&,.wraptext=-1"
      _StyleDefs(75)  =   "Named:id=44:OddRow"
      _StyleDefs(76)  =   ":id=44,.parent=37,.bgcolor=&HD5ECF9&"
      _StyleDefs(77)  =   "Named:id=47:RecordSelector"
      _StyleDefs(78)  =   ":id=47,.parent=38"
      _StyleDefs(79)  =   "Named:id=50:FilterBar"
      _StyleDefs(80)  =   ":id=50,.parent=37"
   End
   Begin Geslab.ControlPanelXP cpDatos 
      Height          =   3960
      Left            =   45
      TabIndex        =   24
      Top             =   405
      Width           =   13830
      _ExtentX        =   24395
      _ExtentY        =   6985
      Caption         =   "Datos de la Oferta"
      BackColor       =   12632256
      HeaderColor     =   8421504
      Object.Height          =   3960
      Begin VB.Frame Frame1 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         ForeColor       =   &H80000008&
         Height          =   3390
         Left            =   135
         TabIndex        =   25
         Top             =   405
         Width           =   13605
         Begin VB.TextBox txtDatos 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00FF80FF&
            Height          =   330
            Index           =   4
            Left            =   11880
            TabIndex        =   5
            Top             =   675
            Width           =   1590
         End
         Begin VB.CheckBox chkENACTexto 
            BackColor       =   &H00C0C0C0&
            Caption         =   "Las actividades marcadas no están amparadas por la acreditación de ENAC"
            ForeColor       =   &H000000FF&
            Height          =   330
            Left            =   6660
            TabIndex        =   51
            Top             =   2970
            Visible         =   0   'False
            Width           =   6855
         End
         Begin VB.TextBox txtDatos 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   690
            Index           =   3
            Left            =   1305
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   8
            Top             =   2250
            Width           =   5955
         End
         Begin VB.TextBox datos 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            DataField       =   "USUARIO"
            DataSource      =   "Adodc1"
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Arial"
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
            Left            =   1305
            Locked          =   -1  'True
            TabIndex        =   40
            Top             =   270
            Width           =   1290
         End
         Begin VB.Frame Frame2 
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            Caption         =   "Tipo Oferta"
            ForeColor       =   &H80000008&
            Height          =   1905
            Left            =   7425
            TabIndex        =   9
            Top             =   1035
            Width           =   1815
            Begin VB.OptionButton opTipo 
               BackColor       =   &H00E0E0E0&
               Caption         =   "Por conceptos"
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
               Index           =   6
               Left            =   90
               TabIndex        =   56
               Top             =   1620
               Width           =   1635
            End
            Begin VB.OptionButton opTipo 
               BackColor       =   &H00E0E0E0&
               Caption         =   "Equipos"
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
               TabIndex        =   53
               Top             =   1395
               Width           =   1635
            End
            Begin VB.OptionButton opTipo 
               BackColor       =   &H00E0E0E0&
               Caption         =   "General"
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
               Index           =   0
               Left            =   90
               TabIndex        =   39
               Top             =   270
               Value           =   -1  'True
               Width           =   1635
            End
            Begin VB.OptionButton opTipo 
               BackColor       =   &H00E0E0E0&
               Caption         =   "Solución"
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
               Left            =   90
               TabIndex        =   38
               Top             =   495
               Width           =   1365
            End
            Begin VB.OptionButton opTipo 
               BackColor       =   &H00E0E0E0&
               Caption         =   "Control Eficacia"
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
               Left            =   90
               TabIndex        =   37
               Top             =   720
               Width           =   1500
            End
            Begin VB.OptionButton opTipo 
               BackColor       =   &H00E0E0E0&
               Caption         =   "Suministro"
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
               Left            =   90
               TabIndex        =   36
               Top             =   945
               Width           =   1635
            End
            Begin VB.OptionButton opTipo 
               BackColor       =   &H00E0E0E0&
               Caption         =   "Agroalimentario"
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
               Index           =   4
               Left            =   90
               TabIndex        =   35
               Top             =   1170
               Width           =   1635
            End
         End
         Begin VB.Frame Frame3 
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            Caption         =   "Logos"
            ForeColor       =   &H80000008&
            Height          =   1770
            Left            =   11340
            TabIndex        =   11
            Top             =   1170
            Width           =   2220
            Begin VB.CheckBox chkLogo 
               BackColor       =   &H00E0E0E0&
               Caption         =   "NADCAP MTL"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   240
               Index           =   5
               Left            =   135
               TabIndex        =   57
               Top             =   1215
               Width           =   1725
            End
            Begin VB.CheckBox chkLogo 
               BackColor       =   &H00E0E0E0&
               Caption         =   "ENAC (Calibraci.)"
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
               Index           =   4
               Left            =   135
               TabIndex        =   52
               Top             =   720
               Width           =   1545
            End
            Begin VB.CheckBox chkLogo 
               BackColor       =   &H00E0E0E0&
               Caption         =   "ENAC (Agroalimenta.)"
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
               Index           =   0
               Left            =   135
               TabIndex        =   34
               Top             =   495
               Width           =   1860
            End
            Begin VB.CheckBox chkLogo 
               BackColor       =   &H00E0E0E0&
               Caption         =   "EQA"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   240
               Index           =   2
               Left            =   135
               TabIndex        =   32
               Top             =   1440
               Width           =   780
            End
            Begin VB.CheckBox chkLogo 
               BackColor       =   &H00E0E0E0&
               Caption         =   "ENAC (Ensayos)"
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
               Index           =   3
               Left            =   135
               TabIndex        =   31
               Top             =   270
               Width           =   1590
            End
            Begin VB.CheckBox chkLogo 
               BackColor       =   &H00E0E0E0&
               Caption         =   "NADCAP"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   240
               Index           =   1
               Left            =   135
               TabIndex        =   33
               Top             =   990
               Width           =   1725
            End
         End
         Begin VB.CheckBox chkSello 
            BackColor       =   &H00C0C0C0&
            Caption         =   "Incluir Sello y Firma"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   5850
            TabIndex        =   1
            Top             =   270
            Value           =   1  'Checked
            Width           =   1815
         End
         Begin VB.TextBox txtDatos 
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
            Height          =   465
            Index           =   0
            Left            =   1305
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   6
            Top             =   1035
            Width           =   5955
         End
         Begin VB.TextBox txtDatos 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   690
            Index           =   1
            Left            =   1305
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   7
            Text            =   "frmOferta_Nueva2.frx":1194
            Top             =   1530
            Width           =   5955
         End
         Begin VB.Frame frameSubtipo 
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            Caption         =   "SubTipo Oferta"
            Enabled         =   0   'False
            ForeColor       =   &H80000008&
            Height          =   1770
            Left            =   9270
            TabIndex        =   10
            Top             =   1170
            Visible         =   0   'False
            Width           =   2040
            Begin VB.OptionButton opsubTipo 
               BackColor       =   &H00E0E0E0&
               Caption         =   "Otros"
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
               Index           =   4
               Left            =   90
               TabIndex        =   30
               Top             =   1350
               Width           =   1635
            End
            Begin VB.OptionButton opsubTipo 
               BackColor       =   &H00E0E0E0&
               Caption         =   "Sol. Químicas Prep."
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
               Left            =   90
               TabIndex        =   29
               Top             =   1035
               Width           =   1815
            End
            Begin VB.OptionButton opsubTipo 
               BackColor       =   &H00E0E0E0&
               Caption         =   "Producto Controlado"
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
               Left            =   90
               TabIndex        =   28
               Top             =   720
               Width           =   1905
            End
            Begin VB.OptionButton opsubTipo 
               BackColor       =   &H00E0E0E0&
               Caption         =   "Alodine"
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
               Left            =   90
               TabIndex        =   27
               Top             =   405
               Width           =   1635
            End
         End
         Begin VB.TextBox datos 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            DataField       =   "USUARIO"
            DataSource      =   "Adodc1"
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   345
            Index           =   5
            Left            =   12240
            Locked          =   -1  'True
            TabIndex        =   26
            Top             =   225
            Width           =   1200
         End
         Begin VB.OptionButton opIdioma 
            BackColor       =   &H00C0C0C0&
            Caption         =   "Oferta en Español"
            Height          =   195
            Index           =   0
            Left            =   135
            TabIndex        =   12
            Top             =   3060
            Value           =   -1  'True
            Width           =   2130
         End
         Begin VB.OptionButton opIdioma 
            BackColor       =   &H00C0C0C0&
            Caption         =   "Oferta en Inglés"
            Height          =   195
            Index           =   1
            Left            =   2250
            TabIndex        =   13
            Top             =   3060
            Width           =   2130
         End
         Begin MSComCtl2.DTPicker fecha 
            Height          =   330
            Left            =   4050
            TabIndex        =   0
            Top             =   270
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
            Format          =   51314689
            CurrentDate     =   38000
            MinDate         =   2
         End
         Begin pryCombo.miCombo cmbclientes 
            Height          =   345
            Left            =   1305
            TabIndex        =   3
            Top             =   675
            Width           =   5955
            _ExtentX        =   10504
            _ExtentY        =   609
         End
         Begin MSDataListLib.DataCombo cmbDatos 
            Height          =   360
            Index           =   1
            Left            =   8280
            TabIndex        =   2
            Top             =   225
            Width           =   3255
            _ExtentX        =   5741
            _ExtentY        =   635
            _Version        =   393216
            Appearance      =   0
            Style           =   2
            Text            =   ""
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin MSComCtl2.DTPicker fechaAceptacion 
            Height          =   330
            Left            =   9405
            TabIndex        =   4
            Top             =   675
            Visible         =   0   'False
            Width           =   1350
            _ExtentX        =   2381
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
            Format          =   51314689
            CurrentDate     =   38000
            MinDate         =   2
         End
         Begin VB.CheckBox chkFechaAceptacion 
            BackColor       =   &H00C0C0C0&
            Caption         =   "Check1"
            Height          =   240
            Left            =   7380
            TabIndex        =   54
            Top             =   720
            Visible         =   0   'False
            Width           =   240
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C0C0&
            Caption         =   "Imp.Pedido (€)"
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
            Index           =   9
            Left            =   10800
            TabIndex        =   55
            Top             =   720
            Width           =   1020
         End
         Begin VB.Label lblFechaAceptacion 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C0C0&
            Caption         =   "F.Aceptación/Rechazo"
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
            Left            =   7650
            TabIndex        =   49
            Top             =   720
            Visible         =   0   'False
            Width           =   1665
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C0C0&
            Caption         =   "Descripción Interna para Canagrosa"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   585
            Index           =   8
            Left            =   135
            TabIndex        =   48
            Top             =   2295
            Width           =   1110
            WordWrap        =   -1  'True
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C0C0&
            Caption         =   "Fecha Oferta"
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
            Index           =   4
            Left            =   2970
            TabIndex        =   47
            Top             =   315
            Width           =   930
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C0C0&
            Caption         =   "Número"
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
            TabIndex        =   46
            Top             =   315
            Width           =   555
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C0C0&
            Caption         =   "Cliente"
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
            Index           =   0
            Left            =   135
            TabIndex        =   45
            Top             =   720
            Width           =   480
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C0C0&
            Caption         =   "Plazo"
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
            Left            =   135
            TabIndex        =   44
            Top             =   1170
            Width           =   390
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C0C0&
            Caption         =   "Observaciones"
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
            Index           =   6
            Left            =   135
            TabIndex        =   43
            Top             =   1755
            Width           =   1065
         End
         Begin VB.Label Label3 
            BackColor       =   &H00C0C0C0&
            Caption         =   "Estado"
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
            Left            =   7695
            TabIndex        =   42
            Top             =   315
            Width           =   600
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C0C0&
            Caption         =   "Edición"
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
            Index           =   7
            Left            =   11655
            TabIndex        =   41
            Top             =   315
            Width           =   525
         End
      End
   End
   Begin VB.CommandButton cmdAdjuntos 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Ficheros Adjuntos"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   915
      Left            =   2700
      Style           =   1  'Graphical
      TabIndex        =   18
      Top             =   8550
      Width           =   1995
   End
   Begin VB.CommandButton cmdCriterio 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Añadir Tipo"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   915
      Left            =   8775
      Picture         =   "frmOferta_Nueva2.frx":11EB
      Style           =   1  'Graphical
      TabIndex        =   17
      Top             =   8550
      Visible         =   0   'False
      Width           =   1275
   End
   Begin VB.TextBox txtDatos 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   2
      Left            =   11700
      TabIndex        =   20
      Top             =   8145
      Width           =   2130
   End
   Begin VB.CommandButton cmdBorrarLinea 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Eliminar Línea"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   915
      Left            =   1395
      Picture         =   "frmOferta_Nueva2.frx":1AB5
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   8550
      Width           =   1275
   End
   Begin VB.CommandButton cmdanadir 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Añadir Línea"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   915
      Left            =   90
      Picture         =   "frmOferta_Nueva2.frx":237F
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   8550
      Width           =   1275
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
      Height          =   915
      Left            =   11700
      Style           =   1  'Graphical
      TabIndex        =   19
      Top             =   8550
      Width           =   1050
   End
   Begin VB.CommandButton cmdcancel 
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
      Height          =   915
      Left            =   12780
      Style           =   1  'Graphical
      TabIndex        =   21
      Top             =   8550
      Width           =   1050
   End
   Begin VB.Label lbltitulo 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "Gestión de Ofertas"
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
      Height          =   360
      Left            =   90
      TabIndex        =   23
      Top             =   0
      Width           =   2610
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      Caption         =   "Total Oferta"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   1
      Left            =   10395
      TabIndex        =   22
      Top             =   8235
      Width           =   1245
   End
   Begin VB.Shape fondo 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   375
      Left            =   0
      Top             =   0
      Width           =   13995
   End
End
Attribute VB_Name = "frmOferta_Nueva2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim tipo As Integer
Public PK As Long
Public PK_EDICION As Integer
Public Nueva_Edicion As Boolean

Dim x As New XArrayDB
Const filas As Integer = 500
Const Col As Integer = 4

Private Enum COLS
    COL1 = 0
    COL2 = 1
    COL3 = 2
    COL4 = 3
End Enum

Private Sub chkFechaAceptacion_Click()
    fechaAceptacion.Enabled = chkFechaAceptacion.Value
    If chkFechaAceptacion.Value = Checked And Format(fechaAceptacion, "yyyy/mm/dd") = "1900/01/01" Then
        fechaAceptacion = Date
    End If
End Sub

Private Sub chkLogo_Click(Index As Integer)
    If Index = 3 Or Index = 4 Then
        If chkLogo(3).Value = Checked Then
            chkENACTexto.visible = True
        Else
            chkENACTexto.visible = False
        End If
    End If
End Sub

Private Sub cmdHistorialCambios_Click()
    'M1108-I
    frmHistorialCambios.PK_TIPO = HC_TIPOS.HC_OFERTAS
    frmHistorialCambios.PK_ID = datos(0)
    frmHistorialCambios.PK_TITULO = "Seguimiento de Oferta Nº" & datos(0)
    frmHistorialCambios.Show 1
    'M1108-F
End Sub

Private Sub cmbDatos_Change(Index As Integer)
    If cmbDatos(1).BoundText = OFERTAS_ESTADOS.OFERTAS_ESTADOS_ACEPTADA Or _
       cmbDatos(1).BoundText = OFERTAS_ESTADOS.OFERTAS_ESTADOS_RECHAZADA Or _
       cmbDatos(1).BoundText = OFERTAS_ESTADOS.OFERTAS_ESTADOS_ANULADA Then
        lblFechaAceptacion.visible = True
        fechaAceptacion.visible = True
        chkFechaAceptacion.visible = True
    Else
        lblFechaAceptacion.visible = False
        fechaAceptacion.visible = False
        chkFechaAceptacion.visible = False
    End If
End Sub

Private Sub cmdAdjuntos_Click()
'M0499-I
    With frmAdjuntos
        .TOBJETO = TOBJETO.TOBJETO_OFERTA
        .COBJETO = PK
        .Show 1
    End With
    Set frmAdjuntos = Nothing
'M0499-F
End Sub

Private Sub cmdAnadir_Click()
    On Error Resume Next
    Dim f As Integer
    Dim c As Integer
    Dim linea As Integer
    linea = grid.Bookmark
    ' Movemos las lineas al final
    For f = filas - 1 To linea Step -1
        For c = 0 To Col
            x(f + 1, c) = x(f, c)
        Next
    Next
    ' Limpiamos donde estamos
    For c = 0 To Col
        x(linea, c) = ""
    Next
    grid.Refresh
    grid.SetFocus
End Sub

Private Sub cmdBorrarLinea_Click()
    On Error Resume Next
    Dim f As Integer
    Dim c As Integer
    Dim linea As Integer
    linea = grid.Bookmark
    ' Movemos las lineas al final
    For f = linea To filas - 1
        For c = 0 To Col
            x(f, c) = x(f + 1, c)
        Next
    Next
    grid.Refresh
    grid.SetFocus
    calcular_total
End Sub

Private Sub cmdcancel_Click()
    PK = 0
    PK_EDICION = 0
    Unload Me
End Sub

Private Sub cmdCriterio_Click()
    frmOferta_Seleccion.TIPO_OFERTA = tipo
    frmOferta_Seleccion.Show 1
    calcular_total
End Sub

Private Sub cmdok_Click()
   On Error GoTo cmdok_Click_Error
    Dim nueva As Boolean
    'M1108-I
    Dim ohc As New clsHistorial_cambios
    'M1108-F
    nueva = False
    If validar = True Then
        If PK <> 0 Then
            If MsgBox("¿Desea generar una nueva edición de la oferta?", vbYesNo + vbQuestion, App.Title) = vbYes Then
                nueva = True
            End If
        Else
            nueva = True
        End If
        Nueva_Edicion = nueva
        Me.MousePointer = 11
        Dim oOferta As New clsOfertas
        Dim oOferta_Detalle As New clsOfertas_detalle
        With oOferta
            If PK = 0 Then
                .setEDICION = 1
            Else
                If nueva Then
                    .setEDICION = datos(5) + 1
                Else
                    .setEDICION = datos(5)
                End If
            End If
'            .setULTIMA = 1
'            .Calcular_Numero
            .setCLIENTE_ID = cmbClientes.getPK_SALIDA
            .setFECHA = Format(fecha.Value, "dd-mm-yyyy")
            .setNUMERO = datos(0)
            If opIdioma(0).Value = True Then
                .setIDIOMA = 0
            Else
                .setIDIOMA = 1
            End If
            If chkLogo(0).Value = Checked Then
                .setLOGO_ENAC = 1
            End If
            If chkLogo(1).Value = Checked Then
                .setLOGO_NADCAP = 1
            End If
            If chkLogo(5).Value = Checked Then
                .setLOGO_NADCAP_MTL = 1
            End If
            If chkLogo(2).Value = Checked Then
                .setLOGO_EQUA = 1
            End If
            If chkLogo(3).Value = Checked Then
                .setLOGO_ENACM = 1
            End If
            If chkLogo(4).Value = Checked Then
                .setLOGO_ENACCAL = 1
            End If
            If chkENACTexto.Value = Checked Then
                .setENAC_TEXTO = 1
            Else
                .setENAC_TEXTO = 0
            End If
            If chkSello.Value = Checked Then
                .setSELLO = 1
            End If
            ' IPEDIDO
            If txtDatos(4) = "" Then
                .setIPEDIDO = "NULL"
            Else
                .setIPEDIDO = moneda_bd(txtDatos(4))
            End If
            .setPLAZO_ENTREGA = txtDatos(0)
            If txtDatos(1) = "" Then
                .setOBSERVACIONES = " "
            Else
                .setOBSERVACIONES = txtDatos(1)
            End If
            .setTIPO_OFERTA = tipo
            .setSUBTIPO_OFERTA = 0
            If tipo = 3 Then
                If opsubTipo(1).Value = True Then
                    .setSUBTIPO_OFERTA = 1
                End If
                If opsubTipo(2).Value = True Then
                    .setSUBTIPO_OFERTA = 2
                End If
                If opsubTipo(3).Value = True Then
                    .setSUBTIPO_OFERTA = 3
                End If
                If opsubTipo(4).Value = True Then
                    .setSUBTIPO_OFERTA = 4
                End If
            End If
            .setTOTAL = txtDatos(2)
            .setUSUARIO_ID = USUARIO.getID_EMPLEADO
            If nueva = False Then
                .setESTADO_OFERTA = cmbDatos(1).BoundText
            Else
                .setESTADO_OFERTA = OFERTAS_ESTADOS.OFERTAS_ESTADOS_GENERADA
            End If
            
            .setDESCRIPCION = txtDatos(3)
            If OFERTAS_ESTADOS.OFERTAS_ESTADOS_ACEPTADA Then
                If chkFechaAceptacion.Value = Checked Then
                    .setFECHA_ACEPTACION = Format(fechaAceptacion, "yyyy-mm-dd")
                Else
                    .setFECHA_ACEPTACION = "1900-01-01"
                End If
            Else
                .setFECHA_ACEPTACION = "1900-01-01"
            End If
    
            If OFERTAS_ESTADOS.OFERTAS_ESTADOS_RECHAZADA Or OFERTAS_ESTADOS.OFERTAS_ESTADOS_ANULADA Then
                If chkFechaAceptacion.Value = Checked Then
                    .setFECHA_ANULACION = Format(fechaAceptacion, "yyyy-mm-dd")
                Else
                    .setFECHA_ANULACION = "1900-01-01"
                End If
            Else
                .setFECHA_ANULACION = "1900-01-01"
            End If
        End With
        Dim OFERTA As Long
        If nueva Then
            'M1108-I
            
            If PK <> 0 Then
                frmMotivo.lbltitulo = "Indique detalladamente el motivo de generación de nueva edición de la Oferta"
                frmMotivo.Show 1
                If Trim(MOTIVO) = "" Then
                    Me.MousePointer = 0
                    MsgBox "Para modificar la Oferta es necesario introducir el motivo de la nueva edición.", vbInformation, App.Title
                    Exit Sub
                End If
            End If
            
            'M1108-F
            OFERTA = oOferta.Insertar
            'M1108-I
            With ohc
                .setTIPO = HC_TIPOS.HC_OFERTAS
                .setIDENTIFICADOR = datos(0)
                .setIDENTIFICADOR_TEXTO = datos(0)
                .setUSUARIO_ID = USUARIO.getID_EMPLEADO
                If PK <> 0 Then
                    .setMOTIVO = "Nueva edición " & datos(5) + 1 & " : " & Trim(MOTIVO)
'                    .setMOTIVO = "Nueva edición " & datos(5) + 1
                Else
                    .setMOTIVO = HC_CREACION
                End If
                .Insertar
            End With
            'M1108-F
        Else
            'M1108-I
'            frmMotivo.lbltitulo = "Indique motivo de modificación (En blanco para no registrar el cambio)"
'            frmMotivo.Show 1
'            If Trim(MOTIVO) = "" Then
'                MsgBox "Para modificar la Oferta es necesario introducir el motivo de la modificación.", vbInformation, App.Title
'                Exit Sub
'            End If
            'M1108-F
            oOferta.Modificar PK, PK_EDICION
            'M1108-I
'            If MOTIVO <> "" Then
'                With ohc
'                    .setTIPO = HC_TIPOS.HC_OFERTAS
'                    .setIDENTIFICADOR = datos(0)
'                    .setIDENTIFICADOR_TEXTO = datos(0)
'                    .setUSUARIO_ID = USUARIO.getID_EMPLEADO
'                    .setMOTIVO = "Modificación de Oferta : " & Trim(MOTIVO)
'                    .Insertar
'                End With
'            End If
            'M1108-F
            oOferta_Detalle.Eliminar (PK)
            OFERTA = PK
        End If
        ' Detalle
        oOferta.Quitar_Ultima oOferta.getNUMERO
        Dim i As Integer
'        Dim bano_anterior As String
        
        For i = x.LowerBound(1) To x.UpperBound(1)
            If Trim(x.Value(i, COLS.COL1)) <> "" Or _
               Trim(x.Value(i, COLS.COL2)) <> "" Or _
               Trim(x.Value(i, COLS.COL3)) <> "" Or _
               Trim(x.Value(i, COLS.COL4)) <> "" Then
                    With oOferta_Detalle
                        .setOFERTA_ID = OFERTA
                        .setEDICION = oOferta.getEDICION
                        .setBANO = CStr(x.Value(i, COLS.COL1))
                        .setDETERMINACION = CStr(x.Value(i, COLS.COL2))
                        .setRANGO = CStr(x.Value(i, COLS.COL3))
                        If CStr(x.Value(i, COLS.COL4)) = "" Then
                            .setPRECIO = ""
                        Else
                            If InStr(1, CStr(x.Value(i, COLS.COL4)), "€") Then
                                .setPRECIO = CStr(x.Value(i, COLS.COL4))
                            Else
                                .setPRECIO = moneda(CStr(x.Value(i, COLS.COL4)))
                            End If
                        End If
'                        If Trim(lista.ListItems(i).Text) = "" Then
'                            .setBANO = bano_anterior
'                        Else
'                            .setBANO = Replace(lista.ListItems(i).Text, vbNewLine, " ")
'                            bano_anterior = Replace(lista.ListItems(i).Text, vbNewLine, " ")
'                        End If
'                        .setDETERMINACION = Replace(lista.ListItems(i).SubItems(1), vbNewLine, " ")
'                        .setRANGO = Replace(lista.ListItems(i).SubItems(2), vbNewLine, " ")
'                        .setPRECIO = lista.ListItems(i).SubItems(3)
                        .setORDEN = i
                        .Insertar
                    End With
            End If
        Next
        ' DUPLICAR LOS ADJUNTOS DE LAS OFERTAS
        ' PK : ID de la edición Origen
        ' OFERTA : ID de la edición Nueva
        If nueva Then '
            If PK <> 0 Then
                Dim oAdjunto As New clsAdjuntos
                oAdjunto.duplicar TOBJETO.TOBJETO_OFERTA, PK, OFERTA
                Set oAdjunto = Nothing
            End If
        End If
        Me.MousePointer = 0
'        MsgBox "La oferta ha sido almacenada correctamente.", vbInformation, App.Title
        Unload Me
    End If

   On Error GoTo 0
   Exit Sub

cmdok_Click_Error:
    Me.MousePointer = 0
    error_grave ("Error " & Err.Number & " (" & Err.Description & ") in procedure cmdok_Click of Formulario frmOferta_Nueva2")
End Sub

Private Sub cmdRequisitos_Click()
    Dim tipo As Integer
    If opTipo(0).Value = True Or opTipo(1).Value = True Or opTipo(2).Value = True Then
        tipo = 1
    ElseIf opTipo(5).Value = True Then
        tipo = 2
    ElseIf opTipo(3).Value = True Then
        tipo = 3
    ElseIf opTipo(4).Value = True Or opTipo(6).Value = True Then
        tipo = 4
    Else
        Exit Sub
    End If
    frmOferta_Requisitos.TIPO_OFERTA = tipo
    frmOferta_Requisitos.PK = PK
    frmOferta_Requisitos.Show 1
End Sub

Private Sub cpDatos_Expand(State As Boolean)
    If State = True Then
'        grid.Top = 4320
'        grid.Height = 3195
        grid.top = 4545
        grid.Height = 3105
    Else
'        grid.Top = 1080
'        grid.Height = 6435
        grid.top = 945
        grid.Height = 6705
    End If
End Sub
Private Sub Form_Load()
    log Me.Name
    cargar_botones Me
    cargar_combo
    inicializar_grid
    opTipo_Click (0)
    fecha = Date
    fechaAceptacion = "1900-01-01"
    chkFechaAceptacion.Value = Unchecked
    Nueva_Edicion = False
    If PK = 0 Then
        Dim oOferta As New clsOfertas
        datos(0) = oOferta.Calcular_Numero
        cmbDatos(1).BoundText = 0
        cmdAdjuntos.Enabled = False
        datos(5) = 1
    Else
        'M1108-I
        cmdHistorialCambios.visible = True
        cmdRequisitos.visible = True
        'M1108-F
        Frame2.Enabled = False
        lbltitulo = "Modificación de Oferta"
'        lbltitulo.BackColor = &H80FF&
        cargar_oferta
        cmdAdjuntos.Enabled = True
    End If
    If USUARIO.getID_EMPLEADO = 7 Then
        chkSello.Enabled = False
    End If
End Sub

Public Sub cargar_combo()
    'Clientes
    llenar_combo cmbClientes, New clsCliente, 0, frmClientes, ""
    Dim oDec As New clsDecodificadora
    oDec.cargar_combo cmbDatos(1), DECODIFICADORA.ESTADOS_OFERTAS
End Sub
Public Sub cargar_oferta()
    If PK > 0 Then
        
        Dim oOferta As New clsOfertas
        With oOferta
            If PK_EDICION = 0 Then
                If .CargaUltima(PK) = False Then
                    MsgBox "Error al cargar la oferta.", vbCritical, App.Title
                    Exit Sub
                End If
                PK_EDICION = .getEDICION
            Else
                If .Carga(PK, PK_EDICION) = False Then
                    MsgBox "Error al cargar la oferta.", vbCritical, App.Title
                    Exit Sub
                End If
            End If
            datos(0) = .getNUMERO
            datos(5) = .getEDICION
            cmbClientes.MostrarElemento .getCLIENTE_ID
            fecha = .getFECHA
            chkSello.Value = .getSELLO
            
            txtDatos(4) = moneda(.getIPEDIDO)
            
            chkLogo(0).Value = .getLOGO_ENAC
            chkLogo(1).Value = .getLOGO_NADCAP
            chkLogo(5).Value = .getLOGO_NADCAP_MTL
            chkLogo(2).Value = .getLOGO_EQUA
            chkLogo(3).Value = .getLOGO_ENACM
            chkLogo(4).Value = .getLOGO_ENACCAL
            
            If .getLOGO_ENACM = 1 Or .getLOGO_ENACCAL Then
                chkENACTexto.visible = True
            End If
            chkENACTexto.Value = .getENAC_TEXTO
            
            opIdioma(.getIDIOMA).Value = True
            opTipo(.getTIPO_OFERTA).Value = True
            If .getSUBTIPO_OFERTA <> 0 Then
                opsubTipo(.getSUBTIPO_OFERTA).Value = True
            End If
            tipo = .getTIPO_OFERTA
            txtDatos(0) = .getPLAZO_ENTREGA
            txtDatos(1) = .getOBSERVACIONES
            txtDatos(2) = .getTOTAL
            cmbDatos(1).BoundText = .getESTADO_OFERTA
            
            txtDatos(3) = .getDESCRIPCION
            If .getFECHA_ACEPTACION <> "01/01/1900" Then
                chkFechaAceptacion.Value = Checked
                fechaAceptacion.Enabled = True
                fechaAceptacion = .getFECHA_ACEPTACION
            Else
                chkFechaAceptacion.Value = Unchecked
                fechaAceptacion.Enabled = False
            End If
            If .getFECHA_ANULACION <> "01/01/1900" Then
                chkFechaAceptacion.Value = Checked
                fechaAceptacion.Enabled = True
                fechaAceptacion = .getFECHA_ANULACION
            Else
                chkFechaAceptacion.Value = Unchecked
                fechaAceptacion.Enabled = False
            End If
        End With
        ' Detalle
        Dim oOferta_Detalle As New clsOfertas_detalle
        Dim rs As ADODB.Recordset
        Set rs = oOferta_Detalle.Listado(PK, PK_EDICION)
        If rs.RecordCount > 0 Then
            Dim bano_ant As String
            Dim BANO As String
            Dim fila As Long
            fila = 0
            Do
                If bano_ant = rs(3) Then
                    BANO = ""
                Else
                    BANO = rs(3)
                    bano_ant = rs(3)
                End If
                ' GRID
                x(fila, COLS.COL1) = CStr(BANO)
                x(fila, COLS.COL2) = CStr(rs(4))
                x(fila, COLS.COL3) = CStr(rs(5))
                x(fila, COLS.COL4) = CStr(rs(6))
                fila = fila + 1
                
                rs.MoveNext
            Loop Until rs.EOF
            grid.Row = 0
            grid.Col = 0
            grid.Refresh
'            calcular_total
        End If
    End If
End Sub

Public Function validar() As Boolean
    If cmbClientes.getTEXTO = "" Then
        MsgBox "Introduzca un cliente para la oferta.", vbCritical, App.Title
        validar = False
        Exit Function
    End If
    Dim i As Integer
    Dim algo As Boolean
    algo = False
    For i = 0 To filas
        If Trim(x.Value(i, COLS.COL1)) <> "" Or _
           Trim(x.Value(i, COLS.COL2)) <> "" Or _
           Trim(x.Value(i, COLS.COL3)) <> "" Or _
           Trim(x.Value(i, COLS.COL4)) <> "" Then
            algo = True
        End If
    Next
    If Not algo Then
        MsgBox "Introduzca algún concepto en la oferta.", vbCritical, App.Title
        validar = False
        Exit Function
    End If
    If opTipo(3).Value = True Then
        If opsubTipo(4).Value = False And _
            opsubTipo(1).Value = False And _
            opsubTipo(2).Value = False And _
            opsubTipo(3).Value = False Then
            MsgBox "Introduzca el subtipo de oferta.", vbCritical, App.Title
            validar = False
            Exit Function
        End If
    End If
            
    If fechaAceptacion.visible = True And Format(fechaAceptacion, "yyyy-mm-dd") = "1900-01-01" Then
        MsgBox "Indique la fecha de aceptación de la oferta.", vbCritical, App.Title
        fechaAceptacion.SetFocus
        validar = False
        Exit Function
    End If
    validar = True
End Function
Private Sub grid_AfterUpdate()
    calcular_total
End Sub

Private Sub grid_KeyPress(KeyAscii As Integer)
    If KeyAscii = 46 And grid.Col = 3 Then
         KeyAscii = 44
    End If
End Sub

Private Sub opIdioma_Click(Index As Integer)
    If Index = 0 Then
        txtDatos(1) = "Pago 30 días Fecha Factura. Ruego faciliten número de cuenta para recibo bancario."
    ElseIf Index = 1 Then
        txtDatos(1) = "Payment: 30 days from Date Invoice. Please provide account number for bank receipt."
    End If
    
End Sub

Private Sub opTipo_Click(Index As Integer)
    tipo = Index
    frameSubtipo.Enabled = False
    frameSubtipo.visible = False
    Select Case Index
    Case 0, 4 ' General, Agroalimentario
        With grid
            .Columns(0).Caption = "Producto"
            .Columns(0).Width = 4800
            .Columns(0).Alignment = dbgLeft
            .Columns(1).Caption = "Parametros"
            .Columns(1).Width = 4700
            .Columns(1).Alignment = dbgLeft
            .Columns(2).Caption = "Procedimiento"
            .Columns(2).Width = 2400
            .Columns(2).Alignment = dbgCenter
            .Columns(2).Locked = False
            .Columns(2).visible = True
            .Columns(3).Caption = "Precio"
            .Columns(3).Width = 700
            .Columns(3).Alignment = dbgRight
        End With
    Case 1 ' Solucion
        With grid
            .Columns(0).Caption = "Tratamiento-Baño"
            .Columns(0).Width = 4800
            .Columns(0).Alignment = dbgLeft
            .Columns(1).Caption = "Determinación"
            .Columns(1).Width = 4700
            .Columns(1).Alignment = dbgLeft
            .Columns(2).Caption = "Rango"
            .Columns(2).Width = 2400
            .Columns(2).Alignment = dbgCenter
            .Columns(2).Locked = False
            .Columns(2).visible = True
            .Columns(3).Caption = "Precio"
            .Columns(3).Width = 700
            .Columns(3).Alignment = dbgRight
        End With
    Case 2 ' CE
        With grid
            .Columns(0).Caption = "Ensayo"
            .Columns(0).Width = 6000
            .Columns(0).Alignment = dbgLeft
            .Columns(1).Caption = "Norma"
            .Columns(1).Width = 5900
            .Columns(1).Alignment = dbgLeft
            .Columns(2).Caption = ""
            .Columns(2).Width = 1800
            .Columns(2).Alignment = dbgCenter
            .Columns(2).Locked = True
            .Columns(2).visible = False
            .Columns(3).Caption = "Precio"
            .Columns(3).Width = 700
            .Columns(3).Alignment = dbgRight
        End With
    Case 3 ' Suministro
        With grid
            .Columns(0).Caption = "Concepto"
            .Columns(0).Width = 6000
            .Columns(0).Alignment = dbgLeft
            .Columns(1).Caption = "Envase/Unidad"
            .Columns(1).Width = 5900
            .Columns(1).Alignment = dbgLeft
            .Columns(2).Caption = ""
            .Columns(2).Width = 1800
            .Columns(2).Alignment = dbgCenter
            .Columns(2).Locked = True
            .Columns(2).visible = False
            .Columns(3).Caption = "Precio"
            .Columns(3).Width = 700
            .Columns(3).Alignment = dbgRight
        End With
        frameSubtipo.Enabled = True
        frameSubtipo.visible = True
    Case 5 'EQUIPOS
        With grid
            .Columns(0).Caption = "Equipo"
            .Columns(0).Width = 4800
            .Columns(0).Alignment = dbgLeft
            .Columns(1).Caption = "Rango"
            .Columns(1).Width = 4700
            .Columns(1).Alignment = dbgCenter
            .Columns(2).Caption = "Procedimiento"
            .Columns(2).Width = 2400
            .Columns(2).Alignment = dbgCenter
            .Columns(2).Locked = False
            .Columns(2).visible = True
            .Columns(3).Caption = "Precio"
            .Columns(3).Width = 700
            .Columns(3).Alignment = dbgRight
        End With
    Case 6 'CONCEPTOS
        With grid
            .Columns(0).Caption = "Concepto"
            .Columns(0).Width = 8800
            .Columns(0).Alignment = dbgLeft
            .Columns(1).Caption = "Cantidad"
            .Columns(1).Width = 1200
            .Columns(1).Alignment = dbgCenter
            .Columns(2).Caption = "Precio"
            .Columns(2).Width = 1500
            .Columns(2).Alignment = dbgRight
            .Columns(2).Locked = False
            .Columns(2).visible = True
            .Columns(3).Caption = "Total"
            .Columns(3).Width = 1500
            .Columns(3).Alignment = dbgRight
        End With
        
    End Select
End Sub
Public Sub calcular_total()
   On Error GoTo calcular_total_Error

    grid.Refresh
    Dim i As Integer
    Dim total As Currency
    For i = 0 To filas
        If Trim(CStr(x(i, COLS.COL4))) <> "" Then
            total = total + CCur(x.Value(i, COLS.COL4))
        End If
    Next
    txtDatos(2) = Format(total, "currency")

   On Error GoTo 0
   Exit Sub

calcular_total_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure calcular_total of Formulario frmOferta_Nueva2"
End Sub
Public Sub inicializar_grid()
    x.ReDim 0, filas, 0, Col
    x.Clear
    Set grid.Array = x
    grid.Refresh
End Sub

Private Sub txtdatos_GotFocus(Index As Integer)
    txtDatos(Index).BackColor = &HC0FFFF
End Sub

Private Sub txtdatos_LostFocus(Index As Integer)
    txtDatos(Index).BackColor = vbWhite
    If Index = 4 Then
        txtDatos(Index) = moneda(txtDatos(Index))
    End If
End Sub
