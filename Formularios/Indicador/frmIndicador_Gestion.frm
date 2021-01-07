VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{77DDF307-D82B-4757-8B3A-106EC9D75D4B}#8.0#0"; "tdbg8.ocx"
Begin VB.Form frmIndicador_Gestion 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Listado de Indicadores"
   ClientHeight    =   9645
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   16860
   Icon            =   "frmIndicador_Gestion.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9645
   ScaleWidth      =   16860
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame3 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Departamentos"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2310
      Left            =   45
      TabIndex        =   11
      Top             =   630
      Width           =   12120
      Begin MSComctlLib.TreeView Tree 
         Height          =   1965
         Left            =   90
         TabIndex        =   12
         Top             =   270
         Width           =   11910
         _ExtentX        =   21008
         _ExtentY        =   3466
         _Version        =   393217
         HideSelection   =   0   'False
         Indentation     =   353
         Style           =   7
         FullRowSelect   =   -1  'True
         ImageList       =   "ImageList1"
         BorderStyle     =   1
         Appearance      =   0
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Información adicional"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1635
      Left            =   12240
      TabIndex        =   9
      Top             =   1305
      Width           =   4560
      Begin VB.TextBox txtComentarios 
         Height          =   1320
         Left            =   90
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   10
         Top             =   225
         Width           =   4380
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Filtro"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   600
      Left            =   12240
      TabIndex        =   5
      Top             =   630
      Width           =   4560
      Begin VB.TextBox txtAnyo 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000002&
         Height          =   300
         Left            =   1845
         MaxLength       =   30
         TabIndex        =   6
         Top             =   180
         Width           =   1245
      End
      Begin MSComCtl2.UpDown UpDownAnyo 
         Height          =   375
         Left            =   3105
         TabIndex        =   7
         Top             =   135
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   661
         _Version        =   393216
         Value           =   7
         Enabled         =   -1  'True
      End
      Begin VB.Label lblCurso 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Año:"
         Height          =   195
         Index           =   2
         Left            =   1350
         TabIndex        =   8
         Top             =   225
         Width           =   345
      End
   End
   Begin VB.CommandButton cmdOk 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Guardar"
      Height          =   870
      Left            =   14715
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   8685
      Width           =   1020
   End
   Begin VB.CommandButton cmdsalir 
      BackColor       =   &H00E0E0E0&
      Caption         =   "ESC-Salir"
      Height          =   870
      Left            =   15765
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   8685
      Width           =   1020
   End
   Begin TrueDBGrid80.TDBGrid grid 
      Height          =   5670
      Left            =   45
      TabIndex        =   0
      Top             =   2970
      Width           =   16755
      _ExtentX        =   29554
      _ExtentY        =   10001
      _LayoutType     =   4
      _RowHeight      =   -2147483647
      _WasPersistedAsPixels=   0
      Columns(0)._VlistStyle=   0
      Columns(0)._MaxComboItems=   5
      Columns(0).Caption=   "ID"
      Columns(0).DataField=   ""
      Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(1)._VlistStyle=   0
      Columns(1)._MaxComboItems=   5
      Columns(1).Caption=   "Descripción"
      Columns(1).DataField=   ""
      Columns(1).NumberFormat=   "General Date"
      Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(2)._VlistStyle=   0
      Columns(2)._MaxComboItems=   5
      Columns(2).Caption=   "Enero"
      Columns(2).DataField=   ""
      Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(3)._VlistStyle=   0
      Columns(3)._MaxComboItems=   5
      Columns(3).Caption=   "Febrero"
      Columns(3).DataField=   ""
      Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(4)._VlistStyle=   0
      Columns(4)._MaxComboItems=   5
      Columns(4).Caption=   "Marzo"
      Columns(4).DataField=   ""
      Columns(4)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(5)._VlistStyle=   0
      Columns(5)._MaxComboItems=   5
      Columns(5).Caption=   "Abril"
      Columns(5).DataField=   ""
      Columns(5)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(6)._VlistStyle=   0
      Columns(6)._MaxComboItems=   5
      Columns(6).Caption=   "Mayo"
      Columns(6).DataField=   ""
      Columns(6)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(7)._VlistStyle=   0
      Columns(7)._MaxComboItems=   5
      Columns(7).Caption=   "Junio"
      Columns(7).DataField=   ""
      Columns(7)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(8)._VlistStyle=   0
      Columns(8)._MaxComboItems=   5
      Columns(8).Caption=   "Julio"
      Columns(8).DataField=   ""
      Columns(8)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(9)._VlistStyle=   0
      Columns(9)._MaxComboItems=   5
      Columns(9).Caption=   "Agosto"
      Columns(9).DataField=   ""
      Columns(9)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(10)._VlistStyle=   0
      Columns(10)._MaxComboItems=   5
      Columns(10).Caption=   "Septiembre"
      Columns(10).DataField=   ""
      Columns(10)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(11)._VlistStyle=   0
      Columns(11)._MaxComboItems=   5
      Columns(11).Caption=   "Octubre"
      Columns(11).DataField=   ""
      Columns(11)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(12)._VlistStyle=   0
      Columns(12)._MaxComboItems=   5
      Columns(12).Caption=   "Noviembre"
      Columns(12).DataField=   ""
      Columns(12)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(13)._VlistStyle=   0
      Columns(13)._MaxComboItems=   5
      Columns(13).Caption=   "Diciembre"
      Columns(13).DataField=   ""
      Columns(13)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(14)._VlistStyle=   0
      Columns(14)._MaxComboItems=   5
      Columns(14).Caption=   "Objetivo"
      Columns(14).DataField=   ""
      Columns(14)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(15)._VlistStyle=   0
      Columns(15)._MaxComboItems=   5
      Columns(15).Caption=   "Valor Ref."
      Columns(15).DataField=   ""
      Columns(15)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(16)._VlistStyle=   0
      Columns(16)._MaxComboItems=   5
      Columns(16).Caption=   "Media"
      Columns(16).DataField=   ""
      Columns(16)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns.Count   =   17
      Splits(0)._UserFlags=   0
      Splits(0).ShowCollapseExpandIcons=   0   'False
      Splits(0).ScrollGroup=   2
      Splits(0).MarqueeStyle=   1
      Splits(0).SizeMode=   2
      Splits(0).AllowRowSizing=   0   'False
      Splits(0).RecordSelectors=   0   'False
      Splits(0).RecordSelectorWidth=   503
      Splits(0).AllowColSelect=   0   'False
      Splits(0).DividerColor=   12632256
      Splits(0).SpringMode=   0   'False
      Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
      Splits(0)._ColumnProps(0)=   "Columns.Count=17"
      Splits(0)._ColumnProps(1)=   "Column(0).Width=2725"
      Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
      Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=2646"
      Splits(0)._ColumnProps(4)=   "Column(0)._EditAlways=0"
      Splits(0)._ColumnProps(5)=   "Column(0).Visible=0"
      Splits(0)._ColumnProps(6)=   "Column(0).Order=1"
      Splits(0)._ColumnProps(7)=   "Column(1).Width=7938"
      Splits(0)._ColumnProps(8)=   "Column(1).DividerColor=0"
      Splits(0)._ColumnProps(9)=   "Column(1)._WidthInPix=7858"
      Splits(0)._ColumnProps(10)=   "Column(1)._EditAlways=0"
      Splits(0)._ColumnProps(11)=   "Column(1)._ColStyle=8192"
      Splits(0)._ColumnProps(12)=   "Column(1).AllowFocus=0"
      Splits(0)._ColumnProps(13)=   "Column(1).Order=2"
      Splits(0)._ColumnProps(14)=   "Column(2).Width=1588"
      Splits(0)._ColumnProps(15)=   "Column(2).DividerColor=0"
      Splits(0)._ColumnProps(16)=   "Column(2)._WidthInPix=1508"
      Splits(0)._ColumnProps(17)=   "Column(2)._EditAlways=0"
      Splits(0)._ColumnProps(18)=   "Column(2)._ColStyle=1"
      Splits(0)._ColumnProps(19)=   "Column(2).Visible=0"
      Splits(0)._ColumnProps(20)=   "Column(2).Order=3"
      Splits(0)._ColumnProps(21)=   "Column(3).Width=1588"
      Splits(0)._ColumnProps(22)=   "Column(3).DividerColor=0"
      Splits(0)._ColumnProps(23)=   "Column(3)._WidthInPix=1508"
      Splits(0)._ColumnProps(24)=   "Column(3)._EditAlways=0"
      Splits(0)._ColumnProps(25)=   "Column(3)._ColStyle=1"
      Splits(0)._ColumnProps(26)=   "Column(3).Visible=0"
      Splits(0)._ColumnProps(27)=   "Column(3).Order=4"
      Splits(0)._ColumnProps(28)=   "Column(4).Width=1588"
      Splits(0)._ColumnProps(29)=   "Column(4).DividerColor=0"
      Splits(0)._ColumnProps(30)=   "Column(4)._WidthInPix=1508"
      Splits(0)._ColumnProps(31)=   "Column(4)._EditAlways=0"
      Splits(0)._ColumnProps(32)=   "Column(4)._ColStyle=1"
      Splits(0)._ColumnProps(33)=   "Column(4).Visible=0"
      Splits(0)._ColumnProps(34)=   "Column(4).Order=5"
      Splits(0)._ColumnProps(35)=   "Column(5).Width=1588"
      Splits(0)._ColumnProps(36)=   "Column(5).DividerColor=0"
      Splits(0)._ColumnProps(37)=   "Column(5)._WidthInPix=1508"
      Splits(0)._ColumnProps(38)=   "Column(5)._EditAlways=0"
      Splits(0)._ColumnProps(39)=   "Column(5)._ColStyle=1"
      Splits(0)._ColumnProps(40)=   "Column(5).Visible=0"
      Splits(0)._ColumnProps(41)=   "Column(5).Order=6"
      Splits(0)._ColumnProps(42)=   "Column(6).Width=1588"
      Splits(0)._ColumnProps(43)=   "Column(6).DividerColor=0"
      Splits(0)._ColumnProps(44)=   "Column(6)._WidthInPix=1508"
      Splits(0)._ColumnProps(45)=   "Column(6)._EditAlways=0"
      Splits(0)._ColumnProps(46)=   "Column(6)._ColStyle=1"
      Splits(0)._ColumnProps(47)=   "Column(6).Visible=0"
      Splits(0)._ColumnProps(48)=   "Column(6).Order=7"
      Splits(0)._ColumnProps(49)=   "Column(7).Width=1588"
      Splits(0)._ColumnProps(50)=   "Column(7).DividerColor=0"
      Splits(0)._ColumnProps(51)=   "Column(7)._WidthInPix=1508"
      Splits(0)._ColumnProps(52)=   "Column(7)._EditAlways=0"
      Splits(0)._ColumnProps(53)=   "Column(7)._ColStyle=1"
      Splits(0)._ColumnProps(54)=   "Column(7).Visible=0"
      Splits(0)._ColumnProps(55)=   "Column(7).Order=8"
      Splits(0)._ColumnProps(56)=   "Column(8).Width=1588"
      Splits(0)._ColumnProps(57)=   "Column(8).DividerColor=0"
      Splits(0)._ColumnProps(58)=   "Column(8)._WidthInPix=1508"
      Splits(0)._ColumnProps(59)=   "Column(8)._EditAlways=0"
      Splits(0)._ColumnProps(60)=   "Column(8)._ColStyle=1"
      Splits(0)._ColumnProps(61)=   "Column(8).Visible=0"
      Splits(0)._ColumnProps(62)=   "Column(8).Order=9"
      Splits(0)._ColumnProps(63)=   "Column(9).Width=1588"
      Splits(0)._ColumnProps(64)=   "Column(9).DividerColor=0"
      Splits(0)._ColumnProps(65)=   "Column(9)._WidthInPix=1508"
      Splits(0)._ColumnProps(66)=   "Column(9)._EditAlways=0"
      Splits(0)._ColumnProps(67)=   "Column(9)._ColStyle=1"
      Splits(0)._ColumnProps(68)=   "Column(9).Visible=0"
      Splits(0)._ColumnProps(69)=   "Column(9).Order=10"
      Splits(0)._ColumnProps(70)=   "Column(10).Width=1588"
      Splits(0)._ColumnProps(71)=   "Column(10).DividerColor=0"
      Splits(0)._ColumnProps(72)=   "Column(10)._WidthInPix=1508"
      Splits(0)._ColumnProps(73)=   "Column(10)._EditAlways=0"
      Splits(0)._ColumnProps(74)=   "Column(10)._ColStyle=1"
      Splits(0)._ColumnProps(75)=   "Column(10).Visible=0"
      Splits(0)._ColumnProps(76)=   "Column(10).Order=11"
      Splits(0)._ColumnProps(77)=   "Column(11).Width=1588"
      Splits(0)._ColumnProps(78)=   "Column(11).DividerColor=0"
      Splits(0)._ColumnProps(79)=   "Column(11)._WidthInPix=1508"
      Splits(0)._ColumnProps(80)=   "Column(11)._EditAlways=0"
      Splits(0)._ColumnProps(81)=   "Column(11)._ColStyle=1"
      Splits(0)._ColumnProps(82)=   "Column(11).Visible=0"
      Splits(0)._ColumnProps(83)=   "Column(11).Order=12"
      Splits(0)._ColumnProps(84)=   "Column(12).Width=1588"
      Splits(0)._ColumnProps(85)=   "Column(12).DividerColor=0"
      Splits(0)._ColumnProps(86)=   "Column(12)._WidthInPix=1508"
      Splits(0)._ColumnProps(87)=   "Column(12)._EditAlways=0"
      Splits(0)._ColumnProps(88)=   "Column(12)._ColStyle=1"
      Splits(0)._ColumnProps(89)=   "Column(12).Visible=0"
      Splits(0)._ColumnProps(90)=   "Column(12).Order=13"
      Splits(0)._ColumnProps(91)=   "Column(13).Width=1588"
      Splits(0)._ColumnProps(92)=   "Column(13).DividerColor=0"
      Splits(0)._ColumnProps(93)=   "Column(13)._WidthInPix=1508"
      Splits(0)._ColumnProps(94)=   "Column(13)._EditAlways=0"
      Splits(0)._ColumnProps(95)=   "Column(13)._ColStyle=1"
      Splits(0)._ColumnProps(96)=   "Column(13).Visible=0"
      Splits(0)._ColumnProps(97)=   "Column(13).Order=14"
      Splits(0)._ColumnProps(98)=   "Column(14).Width=2725"
      Splits(0)._ColumnProps(99)=   "Column(14).DividerColor=0"
      Splits(0)._ColumnProps(100)=   "Column(14)._WidthInPix=2646"
      Splits(0)._ColumnProps(101)=   "Column(14)._EditAlways=0"
      Splits(0)._ColumnProps(102)=   "Column(14).Visible=0"
      Splits(0)._ColumnProps(103)=   "Column(14).Order=15"
      Splits(0)._ColumnProps(104)=   "Column(15).Width=2725"
      Splits(0)._ColumnProps(105)=   "Column(15).DividerColor=0"
      Splits(0)._ColumnProps(106)=   "Column(15)._WidthInPix=2646"
      Splits(0)._ColumnProps(107)=   "Column(15)._EditAlways=0"
      Splits(0)._ColumnProps(108)=   "Column(15).Visible=0"
      Splits(0)._ColumnProps(109)=   "Column(15).Order=16"
      Splits(0)._ColumnProps(110)=   "Column(16).Width=2725"
      Splits(0)._ColumnProps(111)=   "Column(16).DividerColor=0"
      Splits(0)._ColumnProps(112)=   "Column(16)._WidthInPix=2646"
      Splits(0)._ColumnProps(113)=   "Column(16)._EditAlways=0"
      Splits(0)._ColumnProps(114)=   "Column(16).Visible=0"
      Splits(0)._ColumnProps(115)=   "Column(16).Order=17"
      Splits(1)._UserFlags=   0
      Splits(1).ShowCollapseExpandIcons=   0   'False
      Splits(1).ScrollGroup=   2
      Splits(1).MarqueeStyle=   1
      Splits(1).AllowRowSizing=   0   'False
      Splits(1).RecordSelectors=   0   'False
      Splits(1).RecordSelectorWidth=   503
      Splits(1).AllowColSelect=   0   'False
      Splits(1).DividerColor=   12632256
      Splits(1).SpringMode=   0   'False
      Splits(1)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
      Splits(1)._ColumnProps(0)=   "Columns.Count=17"
      Splits(1)._ColumnProps(1)=   "Column(0).Width=2725"
      Splits(1)._ColumnProps(2)=   "Column(0).DividerColor=0"
      Splits(1)._ColumnProps(3)=   "Column(0)._WidthInPix=2646"
      Splits(1)._ColumnProps(4)=   "Column(0)._EditAlways=0"
      Splits(1)._ColumnProps(5)=   "Column(0).Visible=0"
      Splits(1)._ColumnProps(6)=   "Column(0).Order=1"
      Splits(1)._ColumnProps(7)=   "Column(1).Width=8361"
      Splits(1)._ColumnProps(8)=   "Column(1).DividerColor=0"
      Splits(1)._ColumnProps(9)=   "Column(1)._WidthInPix=8281"
      Splits(1)._ColumnProps(10)=   "Column(1)._EditAlways=0"
      Splits(1)._ColumnProps(11)=   "Column(1)._ColStyle=8193"
      Splits(1)._ColumnProps(12)=   "Column(1).Visible=0"
      Splits(1)._ColumnProps(13)=   "Column(1).AllowFocus=0"
      Splits(1)._ColumnProps(14)=   "Column(1).Order=2"
      Splits(1)._ColumnProps(15)=   "Column(2).Width=1402"
      Splits(1)._ColumnProps(16)=   "Column(2).DividerColor=0"
      Splits(1)._ColumnProps(17)=   "Column(2)._WidthInPix=1323"
      Splits(1)._ColumnProps(18)=   "Column(2)._EditAlways=0"
      Splits(1)._ColumnProps(19)=   "Column(2)._ColStyle=1"
      Splits(1)._ColumnProps(20)=   "Column(2).Order=3"
      Splits(1)._ColumnProps(21)=   "Column(3).Width=1402"
      Splits(1)._ColumnProps(22)=   "Column(3).DividerColor=0"
      Splits(1)._ColumnProps(23)=   "Column(3)._WidthInPix=1323"
      Splits(1)._ColumnProps(24)=   "Column(3)._EditAlways=0"
      Splits(1)._ColumnProps(25)=   "Column(3)._ColStyle=1"
      Splits(1)._ColumnProps(26)=   "Column(3).Order=4"
      Splits(1)._ColumnProps(27)=   "Column(4).Width=1402"
      Splits(1)._ColumnProps(28)=   "Column(4).DividerColor=0"
      Splits(1)._ColumnProps(29)=   "Column(4)._WidthInPix=1323"
      Splits(1)._ColumnProps(30)=   "Column(4)._EditAlways=0"
      Splits(1)._ColumnProps(31)=   "Column(4)._ColStyle=1"
      Splits(1)._ColumnProps(32)=   "Column(4).Order=5"
      Splits(1)._ColumnProps(33)=   "Column(5).Width=1402"
      Splits(1)._ColumnProps(34)=   "Column(5).DividerColor=0"
      Splits(1)._ColumnProps(35)=   "Column(5)._WidthInPix=1323"
      Splits(1)._ColumnProps(36)=   "Column(5)._EditAlways=0"
      Splits(1)._ColumnProps(37)=   "Column(5)._ColStyle=1"
      Splits(1)._ColumnProps(38)=   "Column(5).Order=6"
      Splits(1)._ColumnProps(39)=   "Column(6).Width=1402"
      Splits(1)._ColumnProps(40)=   "Column(6).DividerColor=0"
      Splits(1)._ColumnProps(41)=   "Column(6)._WidthInPix=1323"
      Splits(1)._ColumnProps(42)=   "Column(6)._EditAlways=0"
      Splits(1)._ColumnProps(43)=   "Column(6)._ColStyle=1"
      Splits(1)._ColumnProps(44)=   "Column(6).Order=7"
      Splits(1)._ColumnProps(45)=   "Column(7).Width=1402"
      Splits(1)._ColumnProps(46)=   "Column(7).DividerColor=0"
      Splits(1)._ColumnProps(47)=   "Column(7)._WidthInPix=1323"
      Splits(1)._ColumnProps(48)=   "Column(7)._EditAlways=0"
      Splits(1)._ColumnProps(49)=   "Column(7)._ColStyle=1"
      Splits(1)._ColumnProps(50)=   "Column(7).Order=8"
      Splits(1)._ColumnProps(51)=   "Column(8).Width=1402"
      Splits(1)._ColumnProps(52)=   "Column(8).DividerColor=0"
      Splits(1)._ColumnProps(53)=   "Column(8)._WidthInPix=1323"
      Splits(1)._ColumnProps(54)=   "Column(8)._EditAlways=0"
      Splits(1)._ColumnProps(55)=   "Column(8)._ColStyle=1"
      Splits(1)._ColumnProps(56)=   "Column(8).Order=9"
      Splits(1)._ColumnProps(57)=   "Column(9).Width=1402"
      Splits(1)._ColumnProps(58)=   "Column(9).DividerColor=0"
      Splits(1)._ColumnProps(59)=   "Column(9)._WidthInPix=1323"
      Splits(1)._ColumnProps(60)=   "Column(9)._EditAlways=0"
      Splits(1)._ColumnProps(61)=   "Column(9)._ColStyle=1"
      Splits(1)._ColumnProps(62)=   "Column(9).Order=10"
      Splits(1)._ColumnProps(63)=   "Column(10).Width=1402"
      Splits(1)._ColumnProps(64)=   "Column(10).DividerColor=0"
      Splits(1)._ColumnProps(65)=   "Column(10)._WidthInPix=1323"
      Splits(1)._ColumnProps(66)=   "Column(10)._EditAlways=0"
      Splits(1)._ColumnProps(67)=   "Column(10)._ColStyle=1"
      Splits(1)._ColumnProps(68)=   "Column(10).Order=11"
      Splits(1)._ColumnProps(69)=   "Column(11).Width=1402"
      Splits(1)._ColumnProps(70)=   "Column(11).DividerColor=0"
      Splits(1)._ColumnProps(71)=   "Column(11)._WidthInPix=1323"
      Splits(1)._ColumnProps(72)=   "Column(11)._EditAlways=0"
      Splits(1)._ColumnProps(73)=   "Column(11)._ColStyle=1"
      Splits(1)._ColumnProps(74)=   "Column(11).Order=12"
      Splits(1)._ColumnProps(75)=   "Column(12).Width=1402"
      Splits(1)._ColumnProps(76)=   "Column(12).DividerColor=0"
      Splits(1)._ColumnProps(77)=   "Column(12)._WidthInPix=1323"
      Splits(1)._ColumnProps(78)=   "Column(12)._EditAlways=0"
      Splits(1)._ColumnProps(79)=   "Column(12)._ColStyle=1"
      Splits(1)._ColumnProps(80)=   "Column(12).Order=13"
      Splits(1)._ColumnProps(81)=   "Column(13).Width=1402"
      Splits(1)._ColumnProps(82)=   "Column(13).DividerColor=0"
      Splits(1)._ColumnProps(83)=   "Column(13)._WidthInPix=1323"
      Splits(1)._ColumnProps(84)=   "Column(13)._EditAlways=0"
      Splits(1)._ColumnProps(85)=   "Column(13)._ColStyle=1"
      Splits(1)._ColumnProps(86)=   "Column(13).Order=14"
      Splits(1)._ColumnProps(87)=   "Column(14).Width=1402"
      Splits(1)._ColumnProps(88)=   "Column(14).DividerColor=0"
      Splits(1)._ColumnProps(89)=   "Column(14)._WidthInPix=1323"
      Splits(1)._ColumnProps(90)=   "Column(14)._EditAlways=0"
      Splits(1)._ColumnProps(91)=   "Column(14)._ColStyle=1"
      Splits(1)._ColumnProps(92)=   "Column(14).Order=15"
      Splits(1)._ColumnProps(93)=   "Column(15).Width=1402"
      Splits(1)._ColumnProps(94)=   "Column(15).DividerColor=0"
      Splits(1)._ColumnProps(95)=   "Column(15)._WidthInPix=1323"
      Splits(1)._ColumnProps(96)=   "Column(15)._EditAlways=0"
      Splits(1)._ColumnProps(97)=   "Column(15)._ColStyle=1"
      Splits(1)._ColumnProps(98)=   "Column(15).Order=16"
      Splits(1)._ColumnProps(99)=   "Column(16).Width=1402"
      Splits(1)._ColumnProps(100)=   "Column(16).DividerColor=0"
      Splits(1)._ColumnProps(101)=   "Column(16)._WidthInPix=1323"
      Splits(1)._ColumnProps(102)=   "Column(16)._EditAlways=0"
      Splits(1)._ColumnProps(103)=   "Column(16)._ColStyle=8193"
      Splits(1)._ColumnProps(104)=   "Column(16).AllowFocus=0"
      Splits(1)._ColumnProps(105)=   "Column(16).Order=17"
      Splits.Count    =   2
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
      _StyleDefs(36)  =   "Splits(0).Columns(0).Style:id=82,.parent=67"
      _StyleDefs(37)  =   "Splits(0).Columns(0).HeadingStyle:id=79,.parent=68"
      _StyleDefs(38)  =   "Splits(0).Columns(0).FooterStyle:id=80,.parent=69"
      _StyleDefs(39)  =   "Splits(0).Columns(0).EditorStyle:id=81,.parent=71"
      _StyleDefs(40)  =   "Splits(0).Columns(1).Style:id=90,.parent=67,.alignment=0,.bgcolor=&HA9C7F3&"
      _StyleDefs(41)  =   ":id=90,.locked=-1,.bold=0,.fontsize=825,.italic=0,.underline=0,.strikethrough=0"
      _StyleDefs(42)  =   ":id=90,.charset=0"
      _StyleDefs(43)  =   ":id=90,.fontname=MS Sans Serif"
      _StyleDefs(44)  =   "Splits(0).Columns(1).HeadingStyle:id=87,.parent=68"
      _StyleDefs(45)  =   "Splits(0).Columns(1).FooterStyle:id=88,.parent=69"
      _StyleDefs(46)  =   "Splits(0).Columns(1).EditorStyle:id=89,.parent=71,.bold=0,.fontsize=975"
      _StyleDefs(47)  =   ":id=89,.italic=0,.underline=0,.strikethrough=0,.charset=0"
      _StyleDefs(48)  =   ":id=89,.fontname=MS Sans Serif"
      _StyleDefs(49)  =   "Splits(0).Columns(2).Style:id=14,.parent=67,.alignment=2,.bgcolor=&HDEEDFA&"
      _StyleDefs(50)  =   "Splits(0).Columns(2).HeadingStyle:id=11,.parent=68"
      _StyleDefs(51)  =   "Splits(0).Columns(2).FooterStyle:id=12,.parent=69"
      _StyleDefs(52)  =   "Splits(0).Columns(2).EditorStyle:id=13,.parent=71"
      _StyleDefs(53)  =   "Splits(0).Columns(3).Style:id=18,.parent=67,.alignment=2"
      _StyleDefs(54)  =   "Splits(0).Columns(3).HeadingStyle:id=15,.parent=68"
      _StyleDefs(55)  =   "Splits(0).Columns(3).FooterStyle:id=16,.parent=69"
      _StyleDefs(56)  =   "Splits(0).Columns(3).EditorStyle:id=17,.parent=71"
      _StyleDefs(57)  =   "Splits(0).Columns(4).Style:id=22,.parent=67,.alignment=2"
      _StyleDefs(58)  =   "Splits(0).Columns(4).HeadingStyle:id=19,.parent=68"
      _StyleDefs(59)  =   "Splits(0).Columns(4).FooterStyle:id=20,.parent=69"
      _StyleDefs(60)  =   "Splits(0).Columns(4).EditorStyle:id=21,.parent=71"
      _StyleDefs(61)  =   "Splits(0).Columns(5).Style:id=26,.parent=67,.alignment=2"
      _StyleDefs(62)  =   "Splits(0).Columns(5).HeadingStyle:id=23,.parent=68"
      _StyleDefs(63)  =   "Splits(0).Columns(5).FooterStyle:id=24,.parent=69"
      _StyleDefs(64)  =   "Splits(0).Columns(5).EditorStyle:id=25,.parent=71"
      _StyleDefs(65)  =   "Splits(0).Columns(6).Style:id=30,.parent=67,.alignment=2"
      _StyleDefs(66)  =   "Splits(0).Columns(6).HeadingStyle:id=27,.parent=68"
      _StyleDefs(67)  =   "Splits(0).Columns(6).FooterStyle:id=28,.parent=69"
      _StyleDefs(68)  =   "Splits(0).Columns(6).EditorStyle:id=29,.parent=71"
      _StyleDefs(69)  =   "Splits(0).Columns(7).Style:id=34,.parent=67,.alignment=2"
      _StyleDefs(70)  =   "Splits(0).Columns(7).HeadingStyle:id=31,.parent=68"
      _StyleDefs(71)  =   "Splits(0).Columns(7).FooterStyle:id=32,.parent=69"
      _StyleDefs(72)  =   "Splits(0).Columns(7).EditorStyle:id=33,.parent=71"
      _StyleDefs(73)  =   "Splits(0).Columns(8).Style:id=49,.parent=67,.alignment=2"
      _StyleDefs(74)  =   "Splits(0).Columns(8).HeadingStyle:id=35,.parent=68"
      _StyleDefs(75)  =   "Splits(0).Columns(8).FooterStyle:id=36,.parent=69"
      _StyleDefs(76)  =   "Splits(0).Columns(8).EditorStyle:id=46,.parent=71"
      _StyleDefs(77)  =   "Splits(0).Columns(9).Style:id=54,.parent=67,.alignment=2"
      _StyleDefs(78)  =   "Splits(0).Columns(9).HeadingStyle:id=51,.parent=68"
      _StyleDefs(79)  =   "Splits(0).Columns(9).FooterStyle:id=52,.parent=69"
      _StyleDefs(80)  =   "Splits(0).Columns(9).EditorStyle:id=53,.parent=71"
      _StyleDefs(81)  =   "Splits(0).Columns(10).Style:id=58,.parent=67,.alignment=2"
      _StyleDefs(82)  =   "Splits(0).Columns(10).HeadingStyle:id=55,.parent=68"
      _StyleDefs(83)  =   "Splits(0).Columns(10).FooterStyle:id=56,.parent=69"
      _StyleDefs(84)  =   "Splits(0).Columns(10).EditorStyle:id=57,.parent=71"
      _StyleDefs(85)  =   "Splits(0).Columns(11).Style:id=62,.parent=67,.alignment=2"
      _StyleDefs(86)  =   "Splits(0).Columns(11).HeadingStyle:id=59,.parent=68"
      _StyleDefs(87)  =   "Splits(0).Columns(11).FooterStyle:id=60,.parent=69"
      _StyleDefs(88)  =   "Splits(0).Columns(11).EditorStyle:id=61,.parent=71"
      _StyleDefs(89)  =   "Splits(0).Columns(12).Style:id=66,.parent=67,.alignment=2"
      _StyleDefs(90)  =   "Splits(0).Columns(12).HeadingStyle:id=63,.parent=68"
      _StyleDefs(91)  =   "Splits(0).Columns(12).FooterStyle:id=64,.parent=69"
      _StyleDefs(92)  =   "Splits(0).Columns(12).EditorStyle:id=65,.parent=71"
      _StyleDefs(93)  =   "Splits(0).Columns(13).Style:id=86,.parent=67,.alignment=2"
      _StyleDefs(94)  =   "Splits(0).Columns(13).HeadingStyle:id=83,.parent=68"
      _StyleDefs(95)  =   "Splits(0).Columns(13).FooterStyle:id=84,.parent=69"
      _StyleDefs(96)  =   "Splits(0).Columns(13).EditorStyle:id=85,.parent=71"
      _StyleDefs(97)  =   "Splits(0).Columns(14).Style:id=94,.parent=67"
      _StyleDefs(98)  =   "Splits(0).Columns(14).HeadingStyle:id=91,.parent=68"
      _StyleDefs(99)  =   "Splits(0).Columns(14).FooterStyle:id=92,.parent=69"
      _StyleDefs(100) =   "Splits(0).Columns(14).EditorStyle:id=93,.parent=71"
      _StyleDefs(101) =   "Splits(0).Columns(15).Style:id=178,.parent=67"
      _StyleDefs(102) =   "Splits(0).Columns(15).HeadingStyle:id=175,.parent=68"
      _StyleDefs(103) =   "Splits(0).Columns(15).FooterStyle:id=176,.parent=69"
      _StyleDefs(104) =   "Splits(0).Columns(15).EditorStyle:id=177,.parent=71"
      _StyleDefs(105) =   "Splits(0).Columns(16).Style:id=102,.parent=67"
      _StyleDefs(106) =   "Splits(0).Columns(16).HeadingStyle:id=99,.parent=68"
      _StyleDefs(107) =   "Splits(0).Columns(16).FooterStyle:id=100,.parent=69"
      _StyleDefs(108) =   "Splits(0).Columns(16).EditorStyle:id=101,.parent=71"
      _StyleDefs(109) =   "Splits(1).Style:id=103,.parent=1"
      _StyleDefs(110) =   "Splits(1).CaptionStyle:id=112,.parent=4"
      _StyleDefs(111) =   "Splits(1).HeadingStyle:id=104,.parent=2,.namedParent=38"
      _StyleDefs(112) =   "Splits(1).FooterStyle:id=105,.parent=3"
      _StyleDefs(113) =   "Splits(1).InactiveStyle:id=106,.parent=5"
      _StyleDefs(114) =   "Splits(1).SelectedStyle:id=108,.parent=6,.namedParent=40"
      _StyleDefs(115) =   "Splits(1).EditorStyle:id=107,.parent=7,.namedParent=40"
      _StyleDefs(116) =   "Splits(1).HighlightRowStyle:id=109,.parent=8"
      _StyleDefs(117) =   "Splits(1).EvenRowStyle:id=110,.parent=9,.namedParent=43"
      _StyleDefs(118) =   "Splits(1).OddRowStyle:id=111,.parent=10,.namedParent=44"
      _StyleDefs(119) =   "Splits(1).RecordSelectorStyle:id=113,.parent=45"
      _StyleDefs(120) =   "Splits(1).FilterBarStyle:id=114,.parent=48"
      _StyleDefs(121) =   "Splits(1).Columns(0).Style:id=118,.parent=103"
      _StyleDefs(122) =   "Splits(1).Columns(0).HeadingStyle:id=115,.parent=104"
      _StyleDefs(123) =   "Splits(1).Columns(0).FooterStyle:id=116,.parent=105"
      _StyleDefs(124) =   "Splits(1).Columns(0).EditorStyle:id=117,.parent=107"
      _StyleDefs(125) =   "Splits(1).Columns(1).Style:id=122,.parent=103,.alignment=2,.bgcolor=&HA9C7F3&"
      _StyleDefs(126) =   ":id=122,.locked=-1,.bold=0,.fontsize=825,.italic=0,.underline=0,.strikethrough=0"
      _StyleDefs(127) =   ":id=122,.charset=0"
      _StyleDefs(128) =   ":id=122,.fontname=MS Sans Serif"
      _StyleDefs(129) =   "Splits(1).Columns(1).HeadingStyle:id=119,.parent=104"
      _StyleDefs(130) =   "Splits(1).Columns(1).FooterStyle:id=120,.parent=105"
      _StyleDefs(131) =   "Splits(1).Columns(1).EditorStyle:id=121,.parent=107,.bold=0,.fontsize=975"
      _StyleDefs(132) =   ":id=121,.italic=0,.underline=0,.strikethrough=0,.charset=0"
      _StyleDefs(133) =   ":id=121,.fontname=MS Sans Serif"
      _StyleDefs(134) =   "Splits(1).Columns(2).Style:id=126,.parent=103,.alignment=2,.bgcolor=&HDEEDFA&"
      _StyleDefs(135) =   "Splits(1).Columns(2).HeadingStyle:id=123,.parent=104"
      _StyleDefs(136) =   "Splits(1).Columns(2).FooterStyle:id=124,.parent=105"
      _StyleDefs(137) =   "Splits(1).Columns(2).EditorStyle:id=125,.parent=107"
      _StyleDefs(138) =   "Splits(1).Columns(3).Style:id=130,.parent=103,.alignment=2"
      _StyleDefs(139) =   "Splits(1).Columns(3).HeadingStyle:id=127,.parent=104"
      _StyleDefs(140) =   "Splits(1).Columns(3).FooterStyle:id=128,.parent=105"
      _StyleDefs(141) =   "Splits(1).Columns(3).EditorStyle:id=129,.parent=107"
      _StyleDefs(142) =   "Splits(1).Columns(4).Style:id=134,.parent=103,.alignment=2"
      _StyleDefs(143) =   "Splits(1).Columns(4).HeadingStyle:id=131,.parent=104"
      _StyleDefs(144) =   "Splits(1).Columns(4).FooterStyle:id=132,.parent=105"
      _StyleDefs(145) =   "Splits(1).Columns(4).EditorStyle:id=133,.parent=107"
      _StyleDefs(146) =   "Splits(1).Columns(5).Style:id=138,.parent=103,.alignment=2"
      _StyleDefs(147) =   "Splits(1).Columns(5).HeadingStyle:id=135,.parent=104"
      _StyleDefs(148) =   "Splits(1).Columns(5).FooterStyle:id=136,.parent=105"
      _StyleDefs(149) =   "Splits(1).Columns(5).EditorStyle:id=137,.parent=107"
      _StyleDefs(150) =   "Splits(1).Columns(6).Style:id=142,.parent=103,.alignment=2"
      _StyleDefs(151) =   "Splits(1).Columns(6).HeadingStyle:id=139,.parent=104"
      _StyleDefs(152) =   "Splits(1).Columns(6).FooterStyle:id=140,.parent=105"
      _StyleDefs(153) =   "Splits(1).Columns(6).EditorStyle:id=141,.parent=107"
      _StyleDefs(154) =   "Splits(1).Columns(7).Style:id=146,.parent=103,.alignment=2"
      _StyleDefs(155) =   "Splits(1).Columns(7).HeadingStyle:id=143,.parent=104"
      _StyleDefs(156) =   "Splits(1).Columns(7).FooterStyle:id=144,.parent=105"
      _StyleDefs(157) =   "Splits(1).Columns(7).EditorStyle:id=145,.parent=107"
      _StyleDefs(158) =   "Splits(1).Columns(8).Style:id=150,.parent=103,.alignment=2"
      _StyleDefs(159) =   "Splits(1).Columns(8).HeadingStyle:id=147,.parent=104"
      _StyleDefs(160) =   "Splits(1).Columns(8).FooterStyle:id=148,.parent=105"
      _StyleDefs(161) =   "Splits(1).Columns(8).EditorStyle:id=149,.parent=107"
      _StyleDefs(162) =   "Splits(1).Columns(9).Style:id=154,.parent=103,.alignment=2"
      _StyleDefs(163) =   "Splits(1).Columns(9).HeadingStyle:id=151,.parent=104"
      _StyleDefs(164) =   "Splits(1).Columns(9).FooterStyle:id=152,.parent=105"
      _StyleDefs(165) =   "Splits(1).Columns(9).EditorStyle:id=153,.parent=107"
      _StyleDefs(166) =   "Splits(1).Columns(10).Style:id=158,.parent=103,.alignment=2"
      _StyleDefs(167) =   "Splits(1).Columns(10).HeadingStyle:id=155,.parent=104"
      _StyleDefs(168) =   "Splits(1).Columns(10).FooterStyle:id=156,.parent=105"
      _StyleDefs(169) =   "Splits(1).Columns(10).EditorStyle:id=157,.parent=107"
      _StyleDefs(170) =   "Splits(1).Columns(11).Style:id=162,.parent=103,.alignment=2"
      _StyleDefs(171) =   "Splits(1).Columns(11).HeadingStyle:id=159,.parent=104"
      _StyleDefs(172) =   "Splits(1).Columns(11).FooterStyle:id=160,.parent=105"
      _StyleDefs(173) =   "Splits(1).Columns(11).EditorStyle:id=161,.parent=107"
      _StyleDefs(174) =   "Splits(1).Columns(12).Style:id=166,.parent=103,.alignment=2"
      _StyleDefs(175) =   "Splits(1).Columns(12).HeadingStyle:id=163,.parent=104"
      _StyleDefs(176) =   "Splits(1).Columns(12).FooterStyle:id=164,.parent=105"
      _StyleDefs(177) =   "Splits(1).Columns(12).EditorStyle:id=165,.parent=107"
      _StyleDefs(178) =   "Splits(1).Columns(13).Style:id=170,.parent=103,.alignment=2"
      _StyleDefs(179) =   "Splits(1).Columns(13).HeadingStyle:id=167,.parent=104"
      _StyleDefs(180) =   "Splits(1).Columns(13).FooterStyle:id=168,.parent=105"
      _StyleDefs(181) =   "Splits(1).Columns(13).EditorStyle:id=169,.parent=107"
      _StyleDefs(182) =   "Splits(1).Columns(14).Style:id=98,.parent=103,.alignment=2"
      _StyleDefs(183) =   "Splits(1).Columns(14).HeadingStyle:id=95,.parent=104"
      _StyleDefs(184) =   "Splits(1).Columns(14).FooterStyle:id=96,.parent=105"
      _StyleDefs(185) =   "Splits(1).Columns(14).EditorStyle:id=97,.parent=107"
      _StyleDefs(186) =   "Splits(1).Columns(15).Style:id=182,.parent=103,.alignment=2"
      _StyleDefs(187) =   "Splits(1).Columns(15).HeadingStyle:id=179,.parent=104"
      _StyleDefs(188) =   "Splits(1).Columns(15).FooterStyle:id=180,.parent=105"
      _StyleDefs(189) =   "Splits(1).Columns(15).EditorStyle:id=181,.parent=107"
      _StyleDefs(190) =   "Splits(1).Columns(16).Style:id=174,.parent=103,.alignment=2,.bgcolor=&HA9C7F3&"
      _StyleDefs(191) =   ":id=174,.locked=-1"
      _StyleDefs(192) =   "Splits(1).Columns(16).HeadingStyle:id=171,.parent=104"
      _StyleDefs(193) =   "Splits(1).Columns(16).FooterStyle:id=172,.parent=105"
      _StyleDefs(194) =   "Splits(1).Columns(16).EditorStyle:id=173,.parent=107"
      _StyleDefs(195) =   "Named:id=37:Normal"
      _StyleDefs(196) =   ":id=37,.parent=0,.alignment=3,.bgcolor=&H80000018&,.bold=0,.fontsize=825"
      _StyleDefs(197) =   ":id=37,.italic=0,.underline=0,.strikethrough=0,.charset=0"
      _StyleDefs(198) =   ":id=37,.fontname=MS Sans Serif"
      _StyleDefs(199) =   "Named:id=38:Heading"
      _StyleDefs(200) =   ":id=38,.parent=37,.valignment=2,.bgcolor=&HC0C0C0&,.fgcolor=&H80000012&"
      _StyleDefs(201) =   ":id=38,.wraptext=-1"
      _StyleDefs(202) =   "Named:id=39:Footing"
      _StyleDefs(203) =   ":id=39,.parent=37,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(204) =   "Named:id=40:Selected"
      _StyleDefs(205) =   ":id=40,.parent=37,.bgcolor=&H8080FF&,.fgcolor=&H0&,.bold=0,.fontsize=825"
      _StyleDefs(206) =   ":id=40,.italic=0,.underline=0,.strikethrough=0,.charset=0"
      _StyleDefs(207) =   ":id=40,.fontname=MS Sans Serif"
      _StyleDefs(208) =   "Named:id=41:Caption"
      _StyleDefs(209) =   ":id=41,.parent=38,.alignment=2"
      _StyleDefs(210) =   "Named:id=42:HighlightRow"
      _StyleDefs(211) =   ":id=42,.parent=37,.bgcolor=&H80000008&,.fgcolor=&H80000005&"
      _StyleDefs(212) =   "Named:id=43:EvenRow"
      _StyleDefs(213) =   ":id=43,.parent=37,.bgcolor=&HFFFFFF&,.wraptext=-1"
      _StyleDefs(214) =   "Named:id=44:OddRow"
      _StyleDefs(215) =   ":id=44,.parent=37,.bgcolor=&HD5ECF9&"
      _StyleDefs(216) =   "Named:id=47:RecordSelector"
      _StyleDefs(217) =   ":id=47,.parent=38"
      _StyleDefs(218) =   "Named:id=50:FilterBar"
      _StyleDefs(219) =   ":id=50,.parent=37"
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   10755
      Top             =   8865
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   2
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmIndicador_Gestion.frx":1272
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmIndicador_Gestion.frx":24F4
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Image imagen 
      Height          =   480
      Left            =   16335
      Picture         =   "frmIndicador_Gestion.frx":3346
      Top             =   45
      Width           =   480
   End
   Begin VB.Label lbltitulo 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "Listado de Indicadores"
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
      TabIndex        =   3
      Top             =   15
      Width           =   2400
   End
   Begin VB.Label lbltitulo 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "Especifique los datos necesarios para localizar los Indicadores"
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   1
      Left            =   90
      TabIndex        =   2
      Top             =   315
      Width           =   4395
   End
   Begin VB.Shape fondo 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   600
      Left            =   0
      Top             =   0
      Width           =   19275
   End
End
Attribute VB_Name = "frmIndicador_Gestion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim x As New XArrayDB
Private HABILITADO As Boolean
Const filas As Integer = 100
'M1291-I
'Const Col As Integer = 15
Const Col As Integer = 16
'M1291-F
Private Enum COLS
    ID = 0
    DESCRIPCION = 1
    MES01 = 2
    MES02 = 3
    MES03 = 4
    MES04 = 5
    MES05 = 6
    MES06 = 7
    MES07 = 8
    MES08 = 9
    MES09 = 10
    MES10 = 11
    MES11 = 12
    MES12 = 13
    OBJETIVO = 14
'M1291-I
'    media = 15
    VALOR_REFERENCIA = 15
    media = 16
'M1291-F
End Enum
Private Sub inicializar_grid()
    x.ReDim 0, filas, 0, Col
    x.Clear
    Set grid.Array = x
    grid.Refresh
End Sub

Private Sub cmdok_Click()
On Error GoTo fallo
    Dim i As Integer
    Dim j As Integer
'    Dim filas As Integer
    Dim oValores As New clsIndicador_valores
    grid.Refresh
'    filas = grid.VisibleRows
'    If filas = 0 Then
'       filas = 1
'    End If
    Me.MousePointer = 11
'    For i = 0 To grid.VisibleRows - 1   'Filas
    For i = 0 To filas - 1
        If Not IsEmpty(x(i, 0)) And x(i, 0) <> 0 And x(i, 0) <> "" Then
'            oValores.EliminarAnyo x(i, 0), CLng(txtAnyo.Text)
            For j = 2 To Col - 1            'Meses por fila (a partir de la tercera posición partiendo desde cero)
                If Not IsEmpty(x(i, j)) And x(i, j) <> "" Then
                   With oValores
                    .setANNO = CLng(txtAnyo.Text)
                    .setCOMENTARIO = Trim(txtComentarios.Text)
                    .setINDICADOR_ID = x(i, 0)
                    .setMES = CLng(j - 1)
                    .setOBJETIVO = ""
                    .setVALOR = x(i, j)
                    .Insertar
                   End With
                End If
            Next j
        End If
    Next i
    Me.MousePointer = 0
    Set oValores = Nothing
    MsgBox "Se han guardado correctamente los cambios", vbInformation + vbOKOnly, App.Title
    grid.Refresh
    Exit Sub
fallo:
    Me.MousePointer = 0
    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure cmdOk of Formulario frmIndicador_Gestion"
End Sub

Private Sub cmdSalir_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    log (Me.Name)
    cargar_botones Me
    Me.top = 100
    Me.Left = 100
    txtAnyo.Text = CStr(Format(Date, "yyyy"))
    inicializar_grid
    cargar_tree
End Sub
Private Sub cargar_tree()
    Dim nodX As Node
    Tree.Nodes.Clear
    Dim rs As New ADODB.Recordset
    Dim consulta As String
    consulta = "SELECT DISTINCT DEP.VALOR,APA.VALOR,DEP.DESCRIPCION,APA.DESCRIPCION " & _
              "   FROM INDICADOR I " & _
              "   LEFT JOIN decodificadora DEP ON I.DEPARTAMENTO_ID = DEP.VALOR AND DEP.CODIGO = 170 " & _
              "   LEFT JOIN decodificadora APA ON I.APARTADO_ID = APA.VALOR AND APA.CODIGO = 171 " & _
              "  where 1 = 1 " & _
              "  ORDER BY DEP.DESCRIPCION,APA.DESCRIPCION"
    Set rs = datos_bd(consulta)
    Dim departamento As Integer
    Dim clavePadre As String
    Dim claveHijo As String
    departamento = 0
    If rs.RecordCount > 0 Then
        Do
            If departamento <> rs(0) Or departamento = 0 Then
                clavePadre = rs(0) & ";"
                Set nodX = Tree.Nodes.Add(, , clavePadre, rs(2), 1)
                Tree.Nodes(nodX.Index).bold = True
                Tree.Nodes(nodX.Index).Expanded = True
                departamento = rs(0)
            End If
            claveHijo = rs(0) & ";" & rs(1)
            Set nodX = Tree.Nodes.Add(clavePadre, tvwChild, claveHijo, rs(3), 1)
            rs.MoveNext
        Loop Until rs.EOF
    End If
End Sub

Private Sub cargar_lista(departamento As Long, apartado As Long)
    If departamento = 0 Or apartado = 0 Then Exit Sub
    Dim rs As New ADODB.Recordset
    Dim oIndicador As New clsIndicador
    inicializar_grid
    
'    If cmbDepartamento.getTEXTO <> "" Then
'        departamento = cmbDepartamento.getPK_SALIDA
'    End If
'    If cmbApartado.getTEXTO <> "" Then
'        apartado = cmbApartado.getPK_SALIDA
'    End If
    Set rs = oIndicador.Listado(departamento, apartado)
    If rs.RecordCount <> 0 Then
        Dim fila As Long
        fila = 0
        Do
            x(fila, COLS.ID) = CStr(rs(0))
            x(fila, COLS.DESCRIPCION) = CStr(rs(3))
            cargar_rejilla fila, rs(0), CLng(txtAnyo.Text)
            rs.MoveNext
            fila = fila + 1
        Loop Until rs.EOF
        calcularMedias
        grid.Row = 0
        grid.Col = 0
        grid.Refresh
    End If
    Verifica_Perfil (apartado)
    If HABILITADO = False Then
        grid.AllowUpdate = False
    Else
        grid.AllowUpdate = True
    End If
    
    Set oIndicador = Nothing
End Sub

Private Sub cargar_rejilla(fila As Long, ID As Long, ANNO As Long)

'Acceso a indicador_valores
    Dim valores As New clsIndicador_valores
    Dim rsvalores As New ADODB.Recordset
    Set rsvalores = valores.ListadoIndicadorAnno(ID, ANNO)
    
    If rsvalores.RecordCount > 0 Then
        Do
            x(fila, rsvalores("MES") + 1) = CStr(rsvalores("VALOR"))
            rsvalores.MoveNext
        Loop Until rsvalores.EOF
    End If
    
End Sub
Private Sub calcularMedias()
    Dim i As Integer
    Dim j As Integer
    Dim Suma As Single
    Dim meses As Integer
   On Error GoTo calcularMedias_Error

    For i = 0 To filas
        Suma = 0
        meses = 0
        'M1291-I -- Para el cálculo de medias nos fijaremos solo en las columnas que representen valores mensuales
        'For j = 2 To Col
        For j = COLS.MES01 To COLS.MES12
        'M1291-F
            If Not IsEmpty(x(i, j)) And x(i, j) <> "" Then
                If IsNumeric(x(i, j)) Then
                    Suma = Suma + CSng(x(i, j))
                    meses = meses + 1
                End If
            End If
        Next
        If Suma <> 0 And meses <> 0 Then
            x(i, COLS.media) = Format(Suma / meses, "###,###.00")
        End If
    Next

   On Error GoTo 0
   Exit Sub

calcularMedias_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure calcularMedias of Formulario frmIndicador_Gestion"
End Sub
Private Sub grid_AfterColEdit(ByVal ColIndex As Integer)
    calcularMedias
End Sub

Private Sub Tree_Click()
    Dim datos() As String
    If Not IsNull(Tree.Nodes(Tree.selectedItem.Index).Key) Then
        datos = Split(Tree.Nodes(Tree.selectedItem.Index).Key, ";")
        If datos(0) <> "" And datos(1) <> "" Then
         cargar_lista CLng(datos(0)), CLng(datos(1))
        End If
    End If
End Sub

Private Sub UpDownAnyo_DownClick()
    Dim ANYO As String
    Dim nanyo As Long
    
    ANYO = txtAnyo.Text
    nanyo = CInt(txtAnyo.Text)
    nanyo = nanyo - 1
    ANYO = CStr(nanyo)
    txtAnyo.Text = ANYO
    Tree_Click
End Sub

Private Sub UpDownAnyo_UpClick()
    Dim ANYO As String
    Dim nanyo As Long
    
    ANYO = txtAnyo.Text
    nanyo = CInt(txtAnyo.Text)
    nanyo = nanyo + 1
    ANYO = CStr(nanyo)
    txtAnyo.Text = ANYO
    Tree_Click
End Sub

Private Sub Verifica_Perfil(apartado As Long)
        Dim oDecodificadora As New clsDecodificadora
        Dim strDatos As String
        
        HABILITADO = False
        oDecodificadora.Carga_valor DECODIFICADORA.INDICADOR_APARTADOS, apartado
        strDatos = oDecodificadora.getPARAMETROS
        'Obtención de los usuarios de la lista (separados por coma)
        Dim strUsuarios() As String
        Dim intCount As Integer
        Dim VALOR As Integer
                               
        strUsuarios = Split(strDatos, ",")
        
        For intCount = LBound(strUsuarios) To UBound(strUsuarios) 'intcount: número de usuarios
            If Trim(strUsuarios(intCount)) <> "" Then 'Para prevenir el caso de encontrar un , al final de la línea de parámetros
                VALOR = CInt(Solo_Numeros(strUsuarios(intCount)))
                If USUARIO.getID_EMPLEADO = VALOR Then
                   HABILITADO = True
                   Exit Sub
                End If
            End If
        Next intCount
        Set oDecodificadora = Nothing
End Sub

'Public Function Solo_Numeros(ByRef sText As String) As String
  '  Dim sActualChar                 As String * 1
 '   Dim lTotalChar                  As Long
 '   Dim x                           As Long
    
 '   lTotalChar = LenB(sText) \ 2
    
 '   If CBool(lTotalChar) Then
 '       For x = 1 To lTotalChar
 '           sActualChar = Mid$(sText, x, 1)
 '           If IsNumeric(sActualChar) Then Solo_Numeros = Solo_Numeros & sActualChar
 '       Next
 '   End If
    '
'End Function
