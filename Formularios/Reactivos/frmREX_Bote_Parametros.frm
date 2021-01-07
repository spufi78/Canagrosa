VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{77DDF307-D82B-4757-8B3A-106EC9D75D4B}#8.0#0"; "tdbg8.ocx"
Begin VB.Form frmREX_Bote_Parametros 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Bote: Parámetros adicionales"
   ClientHeight    =   6750
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   12990
   Icon            =   "frmREX_Bote_Parametros.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6750
   ScaleWidth      =   12990
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdcancel 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Cancelar"
      Height          =   870
      Left            =   11895
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   5850
      Width           =   1050
   End
   Begin VB.CommandButton cmdok 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Aceptar"
      Height          =   870
      Left            =   10800
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   5850
      Width           =   1050
   End
   Begin Geslab.ControlPanelXP PanelMRC 
      Height          =   4710
      Left            =   0
      TabIndex        =   0
      Top             =   1125
      Width           =   12960
      _ExtentX        =   22860
      _ExtentY        =   8308
      Caption         =   "Requisitos analíticos (M.R ó M.R.C.)"
      BackColor       =   12632256
      TextColor       =   0
      HeaderColor     =   8421504
      PanelOpen       =   0   'False
      Object.Height          =   4710
      Begin VB.Frame Frame7 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Valor e incertidumbre"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2040
         Left            =   135
         TabIndex        =   37
         Top             =   2565
         Width           =   6135
         Begin MSDataListLib.DataCombo cmbIncertidumbre 
            Height          =   315
            Left            =   2520
            TabIndex        =   45
            Top             =   1575
            Width           =   3480
            _ExtentX        =   6138
            _ExtentY        =   556
            _Version        =   393216
            Appearance      =   0
            Text            =   ""
         End
         Begin MSDataListLib.DataCombo cmbProcedimiento 
            Height          =   315
            Left            =   2520
            TabIndex        =   44
            Top             =   1125
            Width           =   3480
            _ExtentX        =   6138
            _ExtentY        =   556
            _Version        =   393216
            Appearance      =   0
            Text            =   ""
         End
         Begin VB.TextBox txtMaxima 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            ForeColor       =   &H00000000&
            Height          =   330
            Left            =   2070
            MaxLength       =   100
            TabIndex        =   8
            Top             =   675
            Width           =   1290
         End
         Begin VB.TextBox txtValor 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            ForeColor       =   &H00000000&
            Height          =   330
            Left            =   2070
            MaxLength       =   100
            TabIndex        =   6
            Top             =   225
            Width           =   1290
         End
         Begin MSDataListLib.DataCombo cmbUnidades 
            Height          =   315
            Left            =   3960
            TabIndex        =   7
            Top             =   270
            Width           =   2040
            _ExtentX        =   3598
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
         Begin MSDataListLib.DataCombo cmbUnidadesIncert 
            Height          =   315
            Left            =   3960
            TabIndex        =   49
            Top             =   675
            Width           =   2040
            _ExtentX        =   3598
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
         Begin VB.Label lblIncertidumbre 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C0C0&
            Caption         =   "Uds.:"
            Height          =   195
            Index           =   5
            Left            =   3510
            TabIndex        =   50
            Top             =   720
            Width           =   375
         End
         Begin VB.Label lblIncertidumbre 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C0C0&
            Caption         =   "Procedimiento de asignación:"
            Height          =   195
            Index           =   4
            Left            =   135
            TabIndex        =   42
            Top             =   1170
            Width           =   2085
         End
         Begin VB.Label lblIncertidumbre 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C0C0&
            Caption         =   "Validez de datos / Incertidumbre: "
            Height          =   195
            Index           =   2
            Left            =   135
            TabIndex        =   41
            Top             =   1620
            Width           =   2385
         End
         Begin VB.Label lblIncertidumbre 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C0C0&
            Caption         =   "Incertidumbre Máxima:"
            Height          =   195
            Index           =   1
            Left            =   135
            TabIndex        =   40
            Top             =   720
            Width           =   1590
         End
         Begin VB.Label lblIncertidumbre 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C0C0&
            Caption         =   "Uds.:"
            Height          =   195
            Index           =   3
            Left            =   3510
            TabIndex        =   39
            Top             =   315
            Width           =   375
         End
         Begin VB.Label lblIncertidumbre 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C0C0&
            Caption         =   "Valor de la propiedad:"
            Height          =   195
            Index           =   0
            Left            =   135
            TabIndex        =   38
            Top             =   315
            Width           =   1545
         End
      End
      Begin VB.Frame Frame3 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Homogeneidad, estabilidad y sistema de producción"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   4110
         Left            =   6390
         TabIndex        =   18
         Top             =   495
         Width           =   6270
         Begin VB.Frame Frame6 
            BackColor       =   &H00C0C0C0&
            Caption         =   "Sistema de producción"
            ForeColor       =   &H80000001&
            Height          =   1230
            Left            =   135
            TabIndex        =   11
            Top             =   2700
            Width           =   5910
            Begin VB.CheckBox chkTipo5 
               Caption         =   "Check1"
               Height          =   195
               Index           =   2
               Left            =   135
               TabIndex        =   32
               Top             =   900
               Width           =   195
            End
            Begin VB.CheckBox chkTipo5 
               Caption         =   "Check1"
               Height          =   195
               Index           =   1
               Left            =   135
               TabIndex        =   31
               Top             =   585
               Width           =   195
            End
            Begin VB.CheckBox chkTipo5 
               Caption         =   "Check1"
               Height          =   195
               Index           =   0
               Left            =   135
               TabIndex        =   30
               Top             =   270
               Width           =   195
            End
            Begin VB.Label lblSistema 
               AutoSize        =   -1  'True
               BackColor       =   &H00C0C0C0&
               Caption         =   "2"
               Height          =   195
               Index           =   2
               Left            =   450
               TabIndex        =   35
               Top             =   900
               Width           =   90
            End
            Begin VB.Label lblSistema 
               AutoSize        =   -1  'True
               BackColor       =   &H00C0C0C0&
               Caption         =   "1"
               Height          =   195
               Index           =   1
               Left            =   450
               TabIndex        =   34
               Top             =   585
               Width           =   90
            End
            Begin VB.Label lblSistema 
               AutoSize        =   -1  'True
               BackColor       =   &H00C0C0C0&
               Caption         =   "0"
               Height          =   195
               Index           =   0
               Left            =   450
               TabIndex        =   33
               Top             =   270
               Width           =   90
            End
         End
         Begin VB.Frame Frame5 
            BackColor       =   &H00C0C0C0&
            Caption         =   "Estabilidad"
            ForeColor       =   &H80000001&
            Height          =   960
            Left            =   135
            TabIndex        =   10
            Top             =   1530
            Width           =   5955
            Begin VB.CheckBox chkTipo4 
               Caption         =   "Check1"
               Height          =   195
               Index           =   1
               Left            =   135
               TabIndex        =   28
               Top             =   630
               Width           =   195
            End
            Begin VB.CheckBox chkTipo4 
               Caption         =   "Check1"
               Height          =   195
               Index           =   0
               Left            =   135
               TabIndex        =   27
               Top             =   270
               Width           =   195
            End
            Begin VB.Label lblEstabilidad 
               AutoSize        =   -1  'True
               BackColor       =   &H00C0C0C0&
               Caption         =   "1"
               Height          =   195
               Index           =   1
               Left            =   495
               TabIndex        =   36
               Top             =   630
               Width           =   90
            End
            Begin VB.Label lblEstabilidad 
               AutoSize        =   -1  'True
               BackColor       =   &H00C0C0C0&
               Caption         =   "0"
               Height          =   195
               Index           =   0
               Left            =   495
               TabIndex        =   29
               Top             =   270
               Width           =   90
            End
         End
         Begin VB.Frame Frame4 
            BackColor       =   &H00C0C0C0&
            Caption         =   "Homogeneidad"
            ForeColor       =   &H80000001&
            Height          =   915
            Left            =   135
            TabIndex        =   9
            Top             =   360
            Width           =   5955
            Begin VB.CheckBox chkTipo3 
               Caption         =   "Check1"
               Height          =   195
               Index           =   0
               Left            =   135
               TabIndex        =   22
               Top             =   270
               Width           =   195
            End
            Begin VB.CheckBox chkTipo3 
               Caption         =   "Check1"
               Height          =   195
               Index           =   2
               Left            =   2970
               TabIndex        =   21
               Top             =   270
               Width           =   195
            End
            Begin VB.CheckBox chkTipo3 
               Caption         =   "Check1"
               Height          =   195
               Index           =   3
               Left            =   2970
               TabIndex        =   20
               Top             =   585
               Width           =   195
            End
            Begin VB.CheckBox chkTipo3 
               Caption         =   "Check1"
               Height          =   195
               Index           =   1
               Left            =   135
               TabIndex        =   19
               Top             =   585
               Width           =   195
            End
            Begin VB.Label lblHomogeneo 
               AutoSize        =   -1  'True
               BackColor       =   &H00C0C0C0&
               Caption         =   "0"
               Height          =   195
               Index           =   0
               Left            =   495
               TabIndex        =   26
               Top             =   270
               Width           =   90
            End
            Begin VB.Label lblHomogeneo 
               AutoSize        =   -1  'True
               BackColor       =   &H00C0C0C0&
               Caption         =   "2"
               Height          =   195
               Index           =   2
               Left            =   3330
               TabIndex        =   25
               Top             =   270
               Width           =   90
            End
            Begin VB.Label lblHomogeneo 
               AutoSize        =   -1  'True
               BackColor       =   &H00C0C0C0&
               Caption         =   "3"
               Height          =   195
               Index           =   3
               Left            =   3330
               TabIndex        =   24
               Top             =   585
               Width           =   90
            End
            Begin VB.Label lblHomogeneo 
               AutoSize        =   -1  'True
               BackColor       =   &H00C0C0C0&
               Caption         =   "1"
               Height          =   195
               Index           =   1
               Left            =   495
               TabIndex        =   23
               Top             =   585
               Width           =   90
            End
         End
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Definición del material"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1995
         Left            =   135
         TabIndex        =   3
         Top             =   495
         Width           =   6135
         Begin MSDataListLib.DataCombo cmbCertificado 
            Height          =   315
            Left            =   2070
            TabIndex        =   43
            Top             =   1125
            Width           =   3930
            _ExtentX        =   6932
            _ExtentY        =   556
            _Version        =   393216
            Appearance      =   0
            Text            =   ""
         End
         Begin VB.TextBox txtTamanyo 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            ForeColor       =   &H00000000&
            Height          =   330
            Left            =   2070
            MaxLength       =   100
            TabIndex        =   5
            Top             =   1530
            Width           =   3945
         End
         Begin VB.TextBox txtAnalito 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            ForeColor       =   &H00000000&
            Height          =   330
            Left            =   2070
            MaxLength       =   100
            TabIndex        =   2
            Top             =   315
            Width           =   3945
         End
         Begin VB.TextBox txtInterferencias 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            ForeColor       =   &H00000000&
            Height          =   330
            Left            =   2070
            MaxLength       =   100
            TabIndex        =   4
            Top             =   720
            Width           =   3945
         End
         Begin VB.Label lblMaterial 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C0C0&
            Caption         =   "Dispone de certificado:"
            Height          =   195
            Index           =   2
            Left            =   135
            TabIndex        =   16
            Top             =   1170
            Width           =   1635
         End
         Begin VB.Label lblMaterial 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C0C0&
            Caption         =   "Tamaño de la muestra:"
            Height          =   195
            Index           =   3
            Left            =   135
            TabIndex        =   17
            Top             =   1575
            Width           =   1620
         End
         Begin VB.Label lblMaterial 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C0C0&
            Caption         =   "Matriz:"
            Height          =   195
            Index           =   1
            Left            =   135
            TabIndex        =   15
            Top             =   765
            Width           =   465
         End
         Begin VB.Label lblMaterial 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C0C0&
            Caption         =   "Mensurando / Analito:"
            Height          =   195
            Index           =   0
            Left            =   135
            TabIndex        =   12
            Top             =   360
            Width           =   1575
         End
      End
   End
   Begin Geslab.ControlPanelXP PanelPC 
      Height          =   4860
      Left            =   0
      TabIndex        =   46
      Top             =   720
      Width           =   12960
      _ExtentX        =   22860
      _ExtentY        =   8573
      Caption         =   "Productos controlados"
      BackColor       =   12632256
      TextColor       =   0
      HeaderColor     =   8421504
      PanelOpen       =   0   'False
      Object.Height          =   4860
      Begin VB.Frame frmParametros 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Parámetros"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   4245
         Left            =   90
         TabIndex        =   47
         Top             =   450
         Width           =   12615
         Begin TrueDBGrid80.TDBGrid grid 
            Height          =   3870
            Left            =   45
            TabIndex        =   48
            Top             =   270
            Width           =   12450
            _ExtentX        =   21960
            _ExtentY        =   6826
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
            Columns(1).Caption=   "Requisitos"
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
            Splits(0)._ColumnProps(1)=   "Column(0).Width=6138"
            Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
            Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=6059"
            Splits(0)._ColumnProps(4)=   "Column(0)._EditAlways=0"
            Splits(0)._ColumnProps(5)=   "Column(0)._ColStyle=1"
            Splits(0)._ColumnProps(6)=   "Column(0).Order=1"
            Splits(0)._ColumnProps(7)=   "Column(1).Width=8070"
            Splits(0)._ColumnProps(8)=   "Column(1).DividerColor=0"
            Splits(0)._ColumnProps(9)=   "Column(1)._WidthInPix=7990"
            Splits(0)._ColumnProps(10)=   "Column(1)._EditAlways=0"
            Splits(0)._ColumnProps(11)=   "Column(1).Order=2"
            Splits(0)._ColumnProps(12)=   "Column(2).Width=5106"
            Splits(0)._ColumnProps(13)=   "Column(2).DividerColor=0"
            Splits(0)._ColumnProps(14)=   "Column(2)._WidthInPix=5027"
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
      End
   End
   Begin VB.Image imagen 
      Height          =   480
      Left            =   12465
      Picture         =   "frmREX_Bote_Parametros.frx":6852
      Top             =   90
      Width           =   480
   End
   Begin VB.Label lbltitulo 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "Reactivos Externos / Productos Controlados: Parámetros adicionales"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   300
      Index           =   0
      Left            =   2475
      TabIndex        =   1
      Top             =   180
      Width           =   8265
   End
   Begin VB.Shape fondo 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   690
      Left            =   -30
      Top             =   0
      Width           =   13080
   End
End
Attribute VB_Name = "frmREX_Bote_Parametros"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'FORMULARIO GENERADOR DE LÍNEAS DE PARÁMETROS
'Variables
Public PK As Long
Dim x As New XArrayDB
Const filas As Integer = 20
Const Col As Integer = 3
Const HOMOGENEIDAD As Integer = 3
Const ESTABILIDAD As Integer = 1
Const PRODUCCION As Integer = 2
Private Enum COLS
    PARAMETRO = 0
    REQUISITO = 1
    VALOR = 2
    unidades = 3
End Enum
Private tipo As Integer

Dim parametros(filas, 5) As String
Dim indice As Integer

Private Sub cmdcancel_Click()
    Unload Me
End Sub

'PPAL. ------------------------------------------
Private Sub Form_Load()
    log (Me.Name)
    inicializar_grid
    cargar_botones Me
    cargarTipo
    cargarCombos
    cargaPrincipal
End Sub
'------------------------------------------------
Private Sub inicializar_grid()
    x.ReDim 0, filas, 0, Col
    x.Clear
    Set grid.Array = x
    grid.Refresh
End Sub
Private Sub cargarTipo()
    Dim oBote As New clsTipos_bote_ex
    oBote.CARGAR PK
    tipo = oBote.getTIPO_M_REFERENCIA_ID
    Set oBote = Nothing

End Sub
Private Sub cargarCombos()
'Combos con valores de descodificadora
    Dim oDeco As New clsDecodificadora
    oDeco.cargar_combo cmbCertificado, DECODIFICADORA.REX_CERTIFICADO
    oDeco.cargar_combo cmbProcedimiento, DECODIFICADORA.REX_PROCEDIMIENTO
    oDeco.cargar_combo cmbIncertidumbre, DECODIFICADORA.REX_INCERTIDUMBRE
    cargar_combo cmbUnidades, New clsUnidades
    'M1332-I
    cargar_combo cmbUnidadesIncert, New clsUnidades
    'M1332-F
End Sub
Private Sub cargaPrincipal()
 Select Case tipo
    Case 1
        PanelMRC.CanExpand = False
        PanelPC.CanExpand = True
        cargarLista
    Case 2 To 3
        PanelMRC.CanExpand = True
        PanelPC.CanExpand = False
        cargarCaptionMR
        cargarFormularioMR
    Case 4 To 8
        PanelMRC.CanExpand = False
        PanelPC.CanExpand = True
        cargarLista
    End Select
End Sub
Private Sub cargarCaptionMR()
'Leyenda en checks (se cargan dinámicamente para que siempre haya correspondencia entre lista de valores y contenidos)
    Dim oDeco As New clsDecodificadora
    Dim i As Long
    'HOMOGENEIDAD
    For i = 0 To HOMOGENEIDAD
        If oDeco.Carga_valor(DECODIFICADORA.REX_HOMOGENEIDAD, i) Then
            lblHomogeneo(i).Caption = oDeco.getDESCRIPCION
        End If
    Next i
    'ESTABILIDAD
    For i = 0 To ESTABILIDAD
        If oDeco.Carga_valor(DECODIFICADORA.REX_ESTABILIDAD, i) Then
            lblEstabilidad(i).Caption = oDeco.getDESCRIPCION
        End If
    Next i
    'PRODUCCIÓN
    For i = 0 To PRODUCCION
        If oDeco.Carga_valor(DECODIFICADORA.REX_SIST_PRODUCCION, i) Then
            lblSistema(i).Caption = oDeco.getDESCRIPCION
        End If
    Next i
    Set oDeco = Nothing
End Sub

Private Sub cargarFormularioMR()
'VALORES M.R. y M.R.C.
    Dim strUso() As String
    Dim intCount As Integer
    Dim oMR As New clsTipos_bote_ex_req_analiticos
    Dim oDeco As New clsDecodificadora
    Dim VALOR As Integer
    
    PanelMRC.CanExpand = True
    PanelMRC.PanelOpen = True
    
    oMR.CargaTipo PK
    'DEFINICIÓN DEL MATERIAL
    txtAnalito = oMR.getANALITO
    txtInterferencias = oMR.getMATRIZ
    cmbCertificado.Text = oMR.getCERTIFICADO
    txtTamanyo = oMR.getTAMANYO
   
    'HOMOGENEIDAD,ESTABILIDAD Y SISTEMA DE PRODUCCIÓN
    '------------------------------------------------
    'HOMOGENEIDAD
    strUso = Split(oMR.getHOMOGENEIDAD, ";")
    For intCount = LBound(strUso) To UBound(strUso)
        If strUso(intCount) <> "" Then 'Para prevenir el caso de encontrar un ; al final de la línea de parámetros
          VALOR = CInt(Solo_Numeros(strUso(intCount)))
          If VALOR <= HOMOGENEIDAD Then
            chkTipo3(VALOR).Value = 1
          End If
        End If
    Next intCount
    'ESTABILIDAD
    strUso = Split(oMR.getESTABILIDAD, ";")
    For intCount = LBound(strUso) To UBound(strUso)
        If strUso(intCount) <> "" Then 'Para prevenir el caso de encontrar un ; al final de la línea de parámetros
          VALOR = CInt(Solo_Numeros(strUso(intCount)))
          If VALOR <= ESTABILIDAD Then
             chkTipo4(VALOR).Value = 1
          End If
        End If
    Next intCount
    'PRODUCCIÓN
    strUso = Split(oMR.getPROCEDIMIENTO, ";")
    For intCount = LBound(strUso) To UBound(strUso)
        If strUso(intCount) <> "" Then 'Para prevenir el caso de encontrar un ; al final de la línea de parámetros
          VALOR = CInt(Solo_Numeros(strUso(intCount)))
          If VALOR <= PRODUCCION Then
             chkTipo5(VALOR).Value = 1
          End If
        End If
    Next intCount
    'VALOR E INCERTIDUMBRE
    '------------------------------------------------
    txtValor.Text = oMR.getVALOR_PROPIEDAD
    cmbUnidades.Text = oMR.getUNIDADES
    'M1332-I
    cmbUnidadesIncert.Text = oMR.getUNIDADES_INCERTIDUMBRE
    'M1332-F
    txtMaxima.Text = oMR.getINCERTIDUMBRE
    cmbProcedimiento.Text = oMR.getPROC_ASIGNACION
    cmbIncertidumbre.Text = oMR.getVALIDEZ
    Set oDeco = Nothing
End Sub

Private Sub cargarLista()
' Parametros
' Común a todos los tipos de reactivos
    Dim rs As ADODB.Recordset
    Dim oTBP As New clsTipos_bote_ex_parametros
    PanelPC.PanelOpen = True
    Set rs = oTBP.Listado(PK)
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
End Sub

Private Sub cmdAnadirPNT_Click()
'Adición a la lista del Documento PNT situado en el combo
    If cmbDocumentos.getPK_SALIDA = 0 Then
        MsgBox "Debe seleccionar un PNT de entre los existentes", vbOK + vbExclamation, "Añadir PNT"
        Exit Sub
    End If
    
    Dim i As Integer
    
    For i = 1 To listaDocumentos.ListItems.Count
        If CLng(listaDocumentos.ListItems(i).Text) = CLng(cmbDocumentos.getPK_SALIDA) Then
            MsgBox "Este PNT ya se encuentra asociado al documento.", vbExclamation, App.Title
            Exit Sub
        End If
    Next
    cmbDocumentos.limpiar
End Sub

Private Sub cmdEliminarPNT_Click()
'Sustracción Documento PNT
    If cmbDocumentos.getPK_SALIDA = 0 Then
        MsgBox "Debe seleccionar un PNT de entre los existentes", vbOK + vbExclamation, "Añadir PNT"
        Exit Sub
    End If
End Sub

Private Sub cmdok_Click()
    Select Case tipo
    Case 1
        insertar_parametros
    Case 2 To 3
        insertar_MR
    Case 4 To 8
        insertar_parametros
    End Select
    Unload Me
End Sub

Private Sub insertar_parametros()
  ' Evidencias
   Dim oTBP As New clsTipos_bote_ex_parametros
   On Error GoTo insertar_parametros_Error
    oTBP.Eliminar PK
    Dim i As Integer
    For i = x.LowerBound(1) To x.UpperBound(1)
        If Trim(x.Value(i, COLS.PARAMETRO)) <> "" Then
            With oTBP
                .setTIPO_BOTE_EX_ID = PK
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
    MsgBox "Los cambios se han guardado correctamente", vbOKOnly + vbInformation, App.Title
   Exit Sub
insertar_parametros_Error:
    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure insertar_parametros of Formulario frmREX_Bote"
End Sub

Private Sub insertar_MR()
'Carga de valores en BBDD
On Error GoTo fallo:
   Dim i As Integer
   Dim valores As String
   Dim oMR As New clsTipos_bote_ex_req_analiticos
   oMR.setTIPOS_BOTE_EX_ID = PK
   valores = ""
   For i = 0 To HOMOGENEIDAD
       If chkTipo3(i).Value = 1 Then
          valores = valores & i & ";"
       End If
   Next i
   oMR.setHOMOGENEIDAD = valores
   valores = ""
   For i = 0 To ESTABILIDAD
       If chkTipo4(i).Value = 1 Then
          valores = valores & i & ";"
       End If
   Next i
   oMR.setESTABILIDAD = valores
   valores = ""
   For i = 0 To PRODUCCION
       If chkTipo5(i).Value = 1 Then
          valores = valores & i & ";"
       End If
   Next i
   oMR.setPROCEDIMIENTO = valores
   'BLOQUE DEFINICIÓN DE MATERIAL
   oMR.setANALITO = Trim(txtAnalito.Text)
   oMR.setMATRIZ = Trim(txtInterferencias.Text)
   oMR.setCERTIFICADO = cmbCertificado.Text
   oMR.setTAMANYO = Trim(txtTamanyo.Text)
   'VALOR E INCERTIDUMBRE
   'M1332-I
   'oMR.setVALOR_PROPIEDAD = IIf(IsNumeric(Trim(txtValor.Text)), txtValor.Text, "")
   oMR.setVALOR_PROPIEDAD = Trim(txtValor.Text)
   'M1332-F
   oMR.setUNIDADES = cmbUnidades.Text
'   oMR.setINCERTIDUMBRE = IIf(IsNumeric(Trim(txtMaxima.Text)), txtMaxima.Text, "")
   oMR.setINCERTIDUMBRE = txtMaxima.Text
   'M1332-I
   oMR.setUNIDADES_INCERTIDUMBRE = cmbUnidadesIncert.Text
   'M1332-F
   oMR.setPROC_ASIGNACION = Trim(cmbProcedimiento.Text)
   oMR.setVALIDEZ = Trim(cmbIncertidumbre.Text)
   oMR.setFECHA_CADUCIDAD = Format(Date, "yyyy-mm-dd")
   oMR.setPROC_ASIGNACION = cmbProcedimiento.Text
   oMR.setVALIDEZ = cmbIncertidumbre.Text
   oMR.EliminarTipo PK
   oMR.Insertar
   
   MsgBox "Los cambios se han guardado correctamente", vbOKOnly + vbInformation, App.Title
   Exit Sub
fallo:
   MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure cmdOk of Formulario frmREX_Bote_Parametros"
End Sub
