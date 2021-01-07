VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{77DDF307-D82B-4757-8B3A-106EC9D75D4B}#8.0#0"; "tdbg8.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#13.2#0"; "Codejock.Controls.v13.2.1.ocx"
Begin VB.Form frmObras_Contratos 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Gestión de Contratos de la obra : "
   ClientHeight    =   10380
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11580
   Icon            =   "frmObras_Contratos.frx":0000
   KeyPreview      =   -1  'True
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   NegotiateMenus  =   0   'False
   ScaleHeight     =   10380
   ScaleWidth      =   11580
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdFinalizar 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Finalizar Contrato"
      Height          =   885
      Left            =   90
      Picture         =   "frmObras_Contratos.frx":08CA
      Style           =   1  'Graphical
      TabIndex        =   43
      ToolTipText     =   "Pulse para dar por finalizado el contrato seleccionado. Sobre este contrato no se acumularan las cantidades suministradas."
      Top             =   9450
      Width           =   1560
   End
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   "Listado de Contratos de la Obra"
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
      Height          =   2175
      Left            =   90
      TabIndex        =   28
      Top             =   405
      Width           =   11400
      Begin VB.TextBox txtdatos 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   3
         Left            =   3645
         ScrollBars      =   2  'Vertical
         TabIndex        =   29
         Top             =   1755
         Width           =   6810
      End
      Begin MSComctlLib.ListView lista 
         Height          =   1440
         Left            =   135
         TabIndex        =   30
         Top             =   270
         Width           =   10350
         _ExtentX        =   18256
         _ExtentY        =   2540
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
      Begin MSComCtl2.DTPicker fecha 
         Height          =   300
         Left            =   765
         TabIndex        =   31
         Top             =   1755
         Width           =   1425
         _ExtentX        =   2514
         _ExtentY        =   529
         _Version        =   393216
         Format          =   51707905
         CurrentDate     =   40679
      End
      Begin XtremeSuiteControls.PushButton cmdAnadirContrato 
         Height          =   660
         Left            =   10575
         TabIndex        =   32
         ToolTipText     =   "Añadir nuevo contrato"
         Top             =   1035
         Width           =   690
         _Version        =   851970
         _ExtentX        =   1217
         _ExtentY        =   1164
         _StockProps     =   79
         Appearance      =   5
         Picture         =   "frmObras_Contratos.frx":1194
      End
      Begin XtremeSuiteControls.PushButton cmdEliminarContrato 
         Height          =   660
         Left            =   10575
         TabIndex        =   33
         ToolTipText     =   "Eliminar contrato y todos sus datos introducidos"
         Top             =   270
         Width           =   690
         _Version        =   851970
         _ExtentX        =   1217
         _ExtentY        =   1164
         _StockProps     =   79
         Appearance      =   5
         Picture         =   "frmObras_Contratos.frx":79F6
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Descripción"
         Height          =   195
         Index           =   3
         Left            =   2520
         TabIndex        =   35
         Top             =   1800
         Width           =   840
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Fecha"
         Height          =   195
         Index           =   5
         Left            =   150
         TabIndex        =   34
         Top             =   1815
         Width           =   450
      End
   End
   Begin VB.CommandButton cmdSalir 
      BackColor       =   &H00E0E0E0&
      Caption         =   "ESC-Salir"
      Height          =   885
      Left            =   10350
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   9450
      Width           =   1155
   End
   Begin XtremeSuiteControls.Resizer resizer 
      Height          =   6795
      Left            =   45
      TabIndex        =   2
      Top             =   2610
      Visible         =   0   'False
      Width           =   11475
      _Version        =   851970
      _ExtentX        =   20241
      _ExtentY        =   11986
      _StockProps     =   1
      VScrollLargeChange=   1500
      VScrollSmallChange=   100
      VScrollMaximum  =   9500
      BorderStyle     =   4
      Begin VB.Frame fmov 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Caption         =   "Documentos Adjuntos al Contrato"
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
         Height          =   2160
         Left            =   45
         TabIndex        =   37
         Top             =   7144
         Width           =   11025
         Begin MSComctlLib.ListView listaAdjuntos 
            Height          =   1440
            Left            =   90
            TabIndex        =   38
            Top             =   270
            Width           =   9990
            _ExtentX        =   17621
            _ExtentY        =   2540
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
         Begin XtremeSuiteControls.PushButton cmdAnadirAdjunto 
            Height          =   660
            Left            =   10125
            TabIndex        =   39
            ToolTipText     =   "Añadir nuevo contrato"
            Top             =   1035
            Width           =   690
            _Version        =   851970
            _ExtentX        =   1217
            _ExtentY        =   1164
            _StockProps     =   79
            Appearance      =   5
            Picture         =   "frmObras_Contratos.frx":E258
         End
         Begin XtremeSuiteControls.PushButton cmdEliminarAdjunto 
            Height          =   660
            Left            =   10125
            TabIndex        =   40
            ToolTipText     =   "Eliminar contrato y todos sus datos introducidos"
            Top             =   270
            Width           =   690
            _Version        =   851970
            _ExtentX        =   1217
            _ExtentY        =   1164
            _StockProps     =   79
            Appearance      =   5
            Picture         =   "frmObras_Contratos.frx":14ABA
         End
         Begin VB.Label Label1 
            BackColor       =   &H00C0C0C0&
            Caption         =   "Doble Click para mostrar el documento"
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
            Left            =   3465
            TabIndex        =   41
            Top             =   1755
            Width           =   3390
         End
      End
      Begin VB.Frame frmContrato 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Caption         =   "Datos del Contrato"
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
         Height          =   7110
         Left            =   45
         TabIndex        =   3
         Top             =   0
         Width           =   11040
         Begin XtremeSuiteControls.GroupBox GroupBox1 
            Height          =   1725
            Left            =   135
            TabIndex        =   4
            Top             =   270
            Width           =   10815
            _Version        =   851970
            _ExtentX        =   19076
            _ExtentY        =   3043
            _StockProps     =   79
            Caption         =   "Datos de Contacto de Administración"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            UseVisualStyle  =   -1  'True
            Begin VB.TextBox txtdatos 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               BeginProperty Font 
                  Name            =   "Microsoft Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   315
               Index           =   0
               Left            =   1395
               TabIndex        =   7
               Top             =   270
               Width           =   9330
            End
            Begin VB.TextBox txtdatos 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               BeginProperty Font 
                  Name            =   "Microsoft Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   315
               Index           =   1
               Left            =   1395
               TabIndex        =   6
               Top             =   630
               Width           =   3300
            End
            Begin VB.TextBox txtdatos 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               BeginProperty Font 
                  Name            =   "Microsoft Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   585
               Index           =   2
               Left            =   1395
               MultiLine       =   -1  'True
               ScrollBars      =   2  'Vertical
               TabIndex        =   5
               Top             =   990
               Width           =   9330
            End
            Begin VB.Label Label7 
               AutoSize        =   -1  'True
               BackColor       =   &H00C0C0C0&
               Caption         =   "Nombre"
               Height          =   195
               Index           =   0
               Left            =   105
               TabIndex        =   10
               Top             =   330
               Width           =   555
            End
            Begin VB.Label Label7 
               AutoSize        =   -1  'True
               BackColor       =   &H00C0C0C0&
               Caption         =   "Teléfono"
               Height          =   195
               Index           =   1
               Left            =   105
               TabIndex        =   9
               Top             =   690
               Width           =   630
            End
            Begin VB.Label Label7 
               AutoSize        =   -1  'True
               BackColor       =   &H00C0C0C0&
               Caption         =   "Observaciones"
               Height          =   195
               Index           =   2
               Left            =   90
               TabIndex        =   8
               Top             =   1185
               Width           =   1065
            End
         End
         Begin XtremeSuiteControls.GroupBox GroupBox2 
            Height          =   1185
            Left            =   135
            TabIndex        =   11
            Top             =   2070
            Width           =   3120
            _Version        =   851970
            _ExtentX        =   5503
            _ExtentY        =   2090
            _StockProps     =   79
            Caption         =   "Fecha de Recepción del Contrato"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            UseVisualStyle  =   -1  'True
            Begin MSComCtl2.DTPicker recepcion_fecha 
               Height          =   345
               Left            =   720
               TabIndex        =   12
               Top             =   270
               Width           =   1425
               _ExtentX        =   2514
               _ExtentY        =   609
               _Version        =   393216
               Format          =   51707905
               CurrentDate     =   40679
            End
            Begin MSDataListLib.DataCombo recepcion_modo 
               Height          =   315
               Left            =   720
               TabIndex        =   13
               Top             =   720
               Width           =   2160
               _ExtentX        =   3810
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
            Begin VB.Label Label7 
               AutoSize        =   -1  'True
               BackColor       =   &H00C0C0C0&
               Caption         =   "Modo"
               Height          =   195
               Index           =   6
               Left            =   135
               TabIndex        =   15
               Top             =   810
               Width           =   405
            End
            Begin VB.Label Label7 
               AutoSize        =   -1  'True
               BackColor       =   &H00C0C0C0&
               Caption         =   "Fecha"
               Height          =   195
               Index           =   4
               Left            =   135
               TabIndex        =   14
               Top             =   375
               Width           =   450
            End
         End
         Begin XtremeSuiteControls.GroupBox GroupBox3 
            Height          =   1320
            Left            =   135
            TabIndex        =   16
            Top             =   3330
            Width           =   3120
            _Version        =   851970
            _ExtentX        =   5503
            _ExtentY        =   2328
            _StockProps     =   79
            Caption         =   "     Fecha de envío BCA"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            UseVisualStyle  =   -1  'True
            Begin VB.CheckBox envio_check 
               BackColor       =   &H00C0C0C0&
               Caption         =   "Check1"
               Height          =   285
               Left            =   135
               TabIndex        =   17
               Top             =   -45
               Width           =   285
            End
            Begin MSComCtl2.DTPicker envio_fecha 
               Height          =   345
               Left            =   720
               TabIndex        =   18
               Top             =   270
               Visible         =   0   'False
               Width           =   1425
               _ExtentX        =   2514
               _ExtentY        =   609
               _Version        =   393216
               Format          =   51707905
               CurrentDate     =   40679
            End
            Begin MSDataListLib.DataCombo envio_modo 
               Height          =   315
               Left            =   720
               TabIndex        =   19
               Top             =   720
               Visible         =   0   'False
               Width           =   2160
               _ExtentX        =   3810
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
            Begin VB.Label envio_lbl1 
               AutoSize        =   -1  'True
               BackColor       =   &H00C0C0C0&
               Caption         =   "Fecha"
               Height          =   195
               Left            =   180
               TabIndex        =   21
               Top             =   315
               Visible         =   0   'False
               Width           =   450
            End
            Begin VB.Label envio_lbl2 
               AutoSize        =   -1  'True
               BackColor       =   &H00C0C0C0&
               Caption         =   "Modo"
               Height          =   195
               Left            =   180
               TabIndex        =   20
               Top             =   810
               Visible         =   0   'False
               Width           =   405
            End
         End
         Begin XtremeSuiteControls.GroupBox GroupBox4 
            Height          =   780
            Left            =   135
            TabIndex        =   22
            Top             =   4725
            Width           =   3120
            _Version        =   851970
            _ExtentX        =   5503
            _ExtentY        =   1376
            _StockProps     =   79
            Caption         =   "     Fecha de adjudicación"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            UseVisualStyle  =   -1  'True
            Begin VB.CheckBox adj_check 
               BackColor       =   &H00C0C0C0&
               Caption         =   "Check1"
               Height          =   285
               Left            =   135
               TabIndex        =   23
               Top             =   -45
               Width           =   285
            End
            Begin MSComCtl2.DTPicker adj_fecha 
               Height          =   345
               Left            =   720
               TabIndex        =   24
               Top             =   270
               Visible         =   0   'False
               Width           =   1425
               _ExtentX        =   2514
               _ExtentY        =   609
               _Version        =   393216
               Format          =   51707905
               CurrentDate     =   40679
            End
            Begin VB.Label adj_label1 
               AutoSize        =   -1  'True
               BackColor       =   &H00C0C0C0&
               Caption         =   "Fecha"
               Height          =   195
               Left            =   135
               TabIndex        =   25
               Top             =   330
               Visible         =   0   'False
               Width           =   450
            End
         End
         Begin TrueDBGrid80.TDBDropDown tArticulos 
            Height          =   4080
            Left            =   3285
            TabIndex        =   26
            Top             =   2700
            Width           =   7230
            _ExtentX        =   12753
            _ExtentY        =   7197
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
            Splits(0).MarqueeStyle=   3
            Splits(0).AllowRowSizing=   0   'False
            Splits(0).RecordSelectors=   0   'False
            Splits(0).RecordSelectorWidth=   503
            Splits(0).AllowColSelect=   0   'False
            Splits(0).DividerColor=   13160660
            Splits(0).SpringMode=   0   'False
            Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
            Splits(0)._ColumnProps(0)=   "Columns.Count=2"
            Splits(0)._ColumnProps(1)=   "Column(0).Width=1693"
            Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
            Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=1588"
            Splits(0)._ColumnProps(4)=   "Column(0)._EditAlways=0"
            Splits(0)._ColumnProps(5)=   "Column(0).Order=1"
            Splits(0)._ColumnProps(6)=   "Column(1).Width=291"
            Splits(0)._ColumnProps(7)=   "Column(1).DividerColor=0"
            Splits(0)._ColumnProps(8)=   "Column(1)._WidthInPix=185"
            Splits(0)._ColumnProps(9)=   "Column(1)._EditAlways=0"
            Splits(0)._ColumnProps(10)=   "Column(1)._ColStyle=0"
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
            _StyleDefs(6)   =   "Style:id=1,.parent=0,.namedParent=33,.bgcolor=&HC0E0FF&,.bold=0,.fontsize=975"
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
            _StyleDefs(42)  =   "Splits(0).Columns(1).Style:id=46,.parent=13,.alignment=0"
            _StyleDefs(43)  =   "Splits(0).Columns(1).HeadingStyle:id=43,.parent=14"
            _StyleDefs(44)  =   "Splits(0).Columns(1).FooterStyle:id=44,.parent=15"
            _StyleDefs(45)  =   "Splits(0).Columns(1).EditorStyle:id=45,.parent=17"
            _StyleDefs(46)  =   "Named:id=33:Normal"
            _StyleDefs(47)  =   ":id=33,.parent=0"
            _StyleDefs(48)  =   "Named:id=34:Heading"
            _StyleDefs(49)  =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
            _StyleDefs(50)  =   ":id=34,.wraptext=-1"
            _StyleDefs(51)  =   "Named:id=35:Footing"
            _StyleDefs(52)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
            _StyleDefs(53)  =   "Named:id=36:Selected"
            _StyleDefs(54)  =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
            _StyleDefs(55)  =   "Named:id=37:Caption"
            _StyleDefs(56)  =   ":id=37,.parent=34,.alignment=2"
            _StyleDefs(57)  =   "Named:id=38:HighlightRow"
            _StyleDefs(58)  =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
            _StyleDefs(59)  =   "Named:id=39:EvenRow"
            _StyleDefs(60)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
            _StyleDefs(61)  =   "Named:id=40:OddRow"
            _StyleDefs(62)  =   ":id=40,.parent=33"
            _StyleDefs(63)  =   "Named:id=41:RecordSelector"
            _StyleDefs(64)  =   ":id=41,.parent=34"
            _StyleDefs(65)  =   "Named:id=42:FilterBar"
            _StyleDefs(66)  =   ":id=42,.parent=33"
         End
         Begin TrueDBGrid80.TDBGrid gridTarifa 
            Height          =   4635
            Left            =   3285
            TabIndex        =   27
            Top             =   2385
            Width           =   7605
            _ExtentX        =   13414
            _ExtentY        =   8176
            _LayoutType     =   4
            _RowHeight      =   -2147483647
            _WasPersistedAsPixels=   0
            Columns(0)._VlistStyle=   0
            Columns(0)._MaxComboItems=   5
            Columns(0).Caption=   "Ref."
            Columns(0).DataField=   ""
            Columns(0).DropDown=   "tArticulos"
            Columns(0).DropDown.vt=   8
            Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(1)._VlistStyle=   0
            Columns(1)._MaxComboItems=   5
            Columns(1).Caption=   "Artículo"
            Columns(1).DataField=   ""
            Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(2)._VlistStyle=   0
            Columns(2)._MaxComboItems=   5
            Columns(2).Caption=   "Cant. Contrato"
            Columns(2).DataField=   ""
            Columns(2).NumberFormat=   "General Number"
            Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(3)._VlistStyle=   0
            Columns(3)._MaxComboItems=   5
            Columns(3).Caption=   "Cant. Sumin"
            Columns(3).DataField=   ""
            Columns(3).NumberFormat=   "General Number"
            Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns.Count   =   4
            Splits(0)._UserFlags=   0
            Splits(0).MarqueeStyle=   1
            Splits(0).AllowRowSizing=   0   'False
            Splits(0).RecordSelectors=   0   'False
            Splits(0).RecordSelectorWidth=   503
            Splits(0).AllowColSelect=   0   'False
            Splits(0).DividerColor=   12632256
            Splits(0).SpringMode=   0   'False
            Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
            Splits(0)._ColumnProps(0)=   "Columns.Count=4"
            Splits(0)._ColumnProps(1)=   "Column(0).Width=1667"
            Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
            Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=1561"
            Splits(0)._ColumnProps(4)=   "Column(0)._EditAlways=0"
            Splits(0)._ColumnProps(5)=   "Column(0)._ColStyle=1"
            Splits(0)._ColumnProps(6)=   "Column(0).Order=1"
            Splits(0)._ColumnProps(7)=   "Column(0).AutoDropDown=1"
            Splits(0)._ColumnProps(8)=   "Column(0).DropDownList=1"
            Splits(0)._ColumnProps(9)=   "Column(0).AutoCompletion=1"
            Splits(0)._ColumnProps(10)=   "Column(1).Width=6641"
            Splits(0)._ColumnProps(11)=   "Column(1).DividerColor=0"
            Splits(0)._ColumnProps(12)=   "Column(1)._WidthInPix=6535"
            Splits(0)._ColumnProps(13)=   "Column(1)._EditAlways=0"
            Splits(0)._ColumnProps(14)=   "Column(1)._ColStyle=8196"
            Splits(0)._ColumnProps(15)=   "Column(1).AllowFocus=0"
            Splits(0)._ColumnProps(16)=   "Column(1).Order=2"
            Splits(0)._ColumnProps(17)=   "Column(2).Width=2355"
            Splits(0)._ColumnProps(18)=   "Column(2).DividerColor=0"
            Splits(0)._ColumnProps(19)=   "Column(2)._WidthInPix=2249"
            Splits(0)._ColumnProps(20)=   "Column(2)._EditAlways=0"
            Splits(0)._ColumnProps(21)=   "Column(2)._ColStyle=2"
            Splits(0)._ColumnProps(22)=   "Column(2).Order=3"
            Splits(0)._ColumnProps(23)=   "Column(3).Width=2196"
            Splits(0)._ColumnProps(24)=   "Column(3).DividerColor=0"
            Splits(0)._ColumnProps(25)=   "Column(3)._WidthInPix=2090"
            Splits(0)._ColumnProps(26)=   "Column(3)._EditAlways=0"
            Splits(0)._ColumnProps(27)=   "Column(3)._ColStyle=8194"
            Splits(0)._ColumnProps(28)=   "Column(3).Order=4"
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
            _StyleDefs(36)  =   "Splits(0).Columns(0).Style:id=24,.parent=11,.alignment=2,.bold=0,.fontsize=975"
            _StyleDefs(37)  =   ":id=24,.italic=0,.underline=0,.strikethrough=0,.charset=0"
            _StyleDefs(38)  =   ":id=24,.fontname=MS Sans Serif"
            _StyleDefs(39)  =   "Splits(0).Columns(0).HeadingStyle:id=21,.parent=12"
            _StyleDefs(40)  =   "Splits(0).Columns(0).FooterStyle:id=22,.parent=13"
            _StyleDefs(41)  =   "Splits(0).Columns(0).EditorStyle:id=23,.parent=15,.bold=0,.fontsize=975"
            _StyleDefs(42)  =   ":id=23,.italic=0,.underline=0,.strikethrough=0,.charset=0"
            _StyleDefs(43)  =   ":id=23,.fontname=MS Sans Serif"
            _StyleDefs(44)  =   "Splits(0).Columns(1).Style:id=28,.parent=11,.locked=-1"
            _StyleDefs(45)  =   "Splits(0).Columns(1).HeadingStyle:id=25,.parent=12"
            _StyleDefs(46)  =   "Splits(0).Columns(1).FooterStyle:id=26,.parent=13"
            _StyleDefs(47)  =   "Splits(0).Columns(1).EditorStyle:id=27,.parent=15"
            _StyleDefs(48)  =   "Splits(0).Columns(2).Style:id=32,.parent=11,.alignment=1"
            _StyleDefs(49)  =   "Splits(0).Columns(2).HeadingStyle:id=29,.parent=12"
            _StyleDefs(50)  =   "Splits(0).Columns(2).FooterStyle:id=30,.parent=13"
            _StyleDefs(51)  =   "Splits(0).Columns(2).EditorStyle:id=31,.parent=15"
            _StyleDefs(52)  =   "Splits(0).Columns(3).Style:id=36,.parent=11,.alignment=1,.locked=-1"
            _StyleDefs(53)  =   "Splits(0).Columns(3).HeadingStyle:id=33,.parent=12"
            _StyleDefs(54)  =   "Splits(0).Columns(3).FooterStyle:id=34,.parent=13"
            _StyleDefs(55)  =   "Splits(0).Columns(3).EditorStyle:id=35,.parent=15"
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
            _StyleDefs(74)  =   ":id=43,.parent=37,.bgcolor=&HFFFFFF&"
            _StyleDefs(75)  =   "Named:id=44:OddRow"
            _StyleDefs(76)  =   ":id=44,.parent=37,.bgcolor=&HE9E9E9&"
            _StyleDefs(77)  =   "Named:id=47:RecordSelector"
            _StyleDefs(78)  =   ":id=47,.parent=38"
            _StyleDefs(79)  =   "Named:id=50:FilterBar"
            _StyleDefs(80)  =   ":id=50,.parent=37"
         End
         Begin XtremeSuiteControls.PushButton cmdGuardarDatosContrato 
            Height          =   660
            Left            =   225
            TabIndex        =   36
            ToolTipText     =   "Añadir nuevo contrato"
            Top             =   5985
            Width           =   2850
            _Version        =   851970
            _ExtentX        =   5027
            _ExtentY        =   1164
            _StockProps     =   79
            Caption         =   "Guardar Datos del Contrato"
            Appearance      =   5
            Picture         =   "frmObras_Contratos.frx":1B31C
         End
         Begin VB.Label Label2 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00C0FFFF&
            Caption         =   "Cantidades Contratadas"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   270
            Left            =   3285
            TabIndex        =   42
            Top             =   2115
            Width           =   7605
         End
      End
   End
   Begin MSComDlg.CommonDialog cd 
      Left            =   8325
      Top             =   9495
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Label lbltitulo 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Caption         =   "Gestión de Contratos de la obra : "
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   345
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   11565
   End
End
Attribute VB_Name = "frmObras_Contratos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public pk As Long

Dim xTarifa As New XArrayDB
Dim xarticulos As New XArrayDB

Const filasTarifa As Integer = 50
Const ColTarifa As Integer = 4
Private Enum ColsTarifa
    ID = 0
    ARTICULO = 1
    CANTIDAD = 2
    CANTIDAD_SUM = 3
End Enum


Private Sub adj_check_Click()
    If adj_check.Value = Checked Then
        adj_fecha.Visible = True
        adj_label1.Visible = True
    Else
        adj_fecha.Visible = False
        adj_label1.Visible = False
    End If
End Sub

Private Sub cmdAnadirAdjunto_Click()
   On Error GoTo cmdAnadirAdjunto_Click_Error

    cd.DialogTitle = "Abrir fichero"
    cd.InitDir = ReadINI(App.Path & "\config.ini", "documentos", "ruta")
    cd.ShowOpen
    If cd.FileName <> "" Then
        Dim s As String
        s = InputBox("Indique la descripcion para el archivo.", App.Title)
        If Trim(s) = "" Then
            MsgBox "Debe indicar una descripción para el archivo.", vbExclamation, App.Title
        Else
            ' Copiar el adjunto
            On Error Resume Next
            Dim origen As String
            Dim destino As String
            origen = cd.FileName
            
            MkDir ReadINI(App.Path & "\config.ini", "documentos", "ruta") & "\CONTRATOS"
            MkDir ReadINI(App.Path & "\config.ini", "documentos", "ruta") & "\CONTRATOS\" & lista.ListItems(lista.SelectedItem.Index)
            destino = ReadINI(App.Path & "\config.ini", "documentos", "ruta") & "\CONTRATOS\" & lista.ListItems(lista.SelectedItem.Index) & "\" & cd.FileTitle
            
            On Error GoTo cmdAnadirAdjunto_Click_Error
            FileCopy origen, destino
            
            ' Insertar en la tabla
            Dim oOCA As New clsObras_contratos_adjuntos
            With oOCA
                .setCONTRATO_ID = lista.ListItems(lista.SelectedItem.Index)
                .setUSUARIO_ID = USUARIO.getID_EMPLEADO
                .setDESCRIPCION = s
                .setDOCUMENTO = cd.FileTitle
                .Insertar
            End With
            cargar_lista_adjuntos lista.ListItems(lista.SelectedItem.Index)
        End If
'        txtmov(5).Text = cd.FileTitle
'        txtmov(1).Text = cd.FileName
    End If

   On Error GoTo 0
   Exit Sub

cmdAnadirAdjunto_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure cmdAnadirAdjunto_Click of Formulario frmObras_Contratos"
End Sub

Private Sub cmdAnadirContrato_Click()
   On Error GoTo cmdAnadirContrato_Click_Error

    If Trim(txtdatos(3)) = "" Then
        MsgBox "Indique la descripción del contrato.", vbExclamation, App.Title
        txtdatos(3).SetFocus
    Else
        Dim oOC As New clsObras_contratos
        With oOC
            .setOBRA_ID = pk
            .setFECHA = Format(fecha, "yyyy-mm-dd")
            .setDESCRIPCION = txtdatos(3)
            .setUSUARIO_ID = USUARIO.getID_EMPLEADO
            .Insertar
        End With
        txtdatos(3) = ""
        fecha = Date
        cargar_lista_contratos
    End If

   On Error GoTo 0
   Exit Sub

cmdAnadirContrato_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure cmdAnadirContrato_Click of Formulario frmObras_Contratos"

End Sub

Private Sub cmdEliminarAdjunto_Click()
    If listaAdjuntos.ListItems.Count > 0 Then
        If MsgBox("¿Esta seguro de eliminar el archivo adjunto seleccionado?", vbYesNo + vbQuestion, App.Title) = vbYes Then
            Dim oOCA As New clsObras_contratos_adjuntos
            oOCA.Eliminar listaAdjuntos.ListItems(listaAdjuntos.SelectedItem.Index)
            cargar_lista_adjuntos lista.ListItems(lista.SelectedItem.Index)
        End If
    End If
End Sub

Private Sub cmdEliminarContrato_Click()
    If lista.ListItems.Count > 0 Then
        If MsgBox("¿Esta seguro de eliminar el contrato seleccionado? Eliminara todos sus datos, archivos adjuntos y demas.", vbYesNo + vbQuestion, App.Title) = vbYes Then
            Dim oOC As New clsObras_contratos
            oOC.Eliminar lista.ListItems(lista.SelectedItem.Index)
            cargar_lista_contratos
        End If
    End If
End Sub

Private Sub cmdFinalizar_Click()
    If lista.ListItems.Count > 0 Then
        If MsgBox("¿Desea dar por finalizado el contrato marcado? Las cantidades que se albaranen ya no se acumularan en dicho contrato.", vbExclamation + vbYesNo, App.Title) = vbYes Then
            Dim oOC As New clsObras_contratos
            oOC.Modificar_Finalizado lista.ListItems(lista.SelectedItem.Index)
            Set oOC = Nothing
            cargar_lista_contratos
        End If
    End If
End Sub

Private Sub cmdGuardarDatosContrato_Click()
    Dim oOCD As New clsObras_contratos_detalle
   On Error GoTo cmdGuardarDatosContrato_Click_Error

    With oOCD
        .setCONTRATO_ID = lista.ListItems(lista.SelectedItem.Index).Text
        .setNOMBRE = txtdatos(0)
        .setTELEFONO = txtdatos(1)
        .setOBSERVACIONES = txtdatos(2)
        ' Recepcion
        .setRECEPCION_FECHA = Format(recepcion_fecha, "yyyy-mm-dd")
        If recepcion_modo.Text = "" Then
            .setRECEPCION_MODO = 0
        Else
            .setRECEPCION_MODO = recepcion_modo.BoundText
        End If
        ' Envio
        If envio_check = vbChecked Then
            .setENVIO_FECHA = Format(envio_fecha, "yyyy-mm-dd")
            If envio_modo.Text = "" Then
                .setENVIO_MODO = 0
            Else
                .setENVIO_MODO = envio_modo.BoundText
            End If
        Else
            .setENVIO_FECHA = "0000-00-00"
            .setENVIO_MODO = 0
        End If
        ' Adjudicacion
        If adj_check = vbChecked Then
            .setADJUDICACION_FECHA = Format(adj_fecha, "yyyy-mm-dd")
        Else
            .setADJUDICACION_FECHA = "0000-00-00"
        End If
        .Insertar
    End With
    ' Cantidades Contratadas
    Dim oOCC As New clsObras_contratos_cantidades
    Dim i As Integer
    For i = xTarifa.LowerBound(1) To xTarifa.UpperBound(1)
        If Trim(xTarifa.Value(i, ColsTarifa.ID)) <> "" Then
            With oOCC
                .setCONTRATO_ID = lista.ListItems(lista.SelectedItem.Index).Text
                .setARTICULO_ID = Trim(xTarifa.Value(i, ColsTarifa.ID))
                If Trim(xTarifa.Value(i, ColsTarifa.CANTIDAD)) = "" Then
                    .setCANTIDAD_CONTRATADA = "0"
                Else
                    .setCANTIDAD_CONTRATADA = xTarifa.Value(i, ColsTarifa.CANTIDAD)
                End If
                If Trim(xTarifa.Value(i, ColsTarifa.CANTIDAD_SUM)) = "" Then
                    .setCANTIDAD_SUMINISTRADA = "0"
                Else
                    .setCANTIDAD_SUMINISTRADA = xTarifa.Value(i, ColsTarifa.CANTIDAD_SUM)
                End If
                If .Insertar = 0 Then
                    MsgBox "Error al almacenar las cantidades contratadas.", vbExclamation, App.Title
                    Exit Sub
                End If
            End With
        End If
    Next
    
    MsgBox "Los datos del contrato se han almacenado correctamente.", vbInformation, App.Title
    Set oOCD = Nothing

   On Error GoTo 0
   Exit Sub

cmdGuardarDatosContrato_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure cmdGuardarDatosContrato_Click of Formulario frmObras_Contratos"
End Sub

Private Sub cmdSalir_Click()
    Unload Me
End Sub

Private Sub envio_check_Click()
    If envio_check.Value = vbChecked Then
        envio_fecha.Visible = True
        envio_modo.Visible = True
        envio_lbl1.Visible = True
        envio_lbl2.Visible = True
    Else
        envio_fecha.Visible = False
        envio_modo.Visible = False
        envio_lbl1.Visible = False
        envio_lbl2.Visible = False
    End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case 27 ' ESC
            cmdSalir_Click
    End Select
End Sub
Private Sub Form_Load()
    log Me.Name
    cargar_botones Me
    cabecera
    inicializar_grid
    cargar_combos
    If pk > 0 Then
        Dim oObra As New clsObras
        oObra.Carga pk
        Me.Caption = Me.Caption & oObra.getNOMBRE
        lbltitulo = Me.Caption
        Set oObra = Nothing
        cargar_lista_contratos
    End If
End Sub
Private Sub inicializar_grid()
   On Error GoTo inicializar_grid_Error
   
    gridTarifa.Col = 0
    gridTarifa.Row = 0
    xTarifa.Clear
    xTarifa.ReDim 0, filasTarifa, 0, ColTarifa
    xTarifa.Clear
    Set gridTarifa.Array = xTarifa
    gridTarifa.Refresh

    cargar_articulos

   On Error GoTo 0
   Exit Sub

inicializar_grid_Error:
    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure inicializar_grid of Formulario frmCE_Recepcion_Nuevo2"
End Sub

Private Sub cargar_lista_contratos()
    On Error GoTo fallo
    Dim oOC As New clsObras_contratos
    Dim rs As ADODB.Recordset
    Set rs = oOC.Listado(pk)
    resizer.Visible = False
    lista.ListItems.Clear
    If rs.RecordCount > 0 Then
        Do
            With lista.ListItems.Add(, , rs(0))
                .SubItems(1) = Format(rs(1), "DD-MM-YYYY")
                .SubItems(2) = rs(2)
                If rs(3) = 0 Then
                    .SubItems(3) = "EN VIGOR"
                Else
                    .SubItems(3) = "FINALIZADO"
                End If
                .SubItems(4) = rs(4)
            End With
            rs.MoveNext
        Loop Until rs.EOF
        cargar_contrato lista.ListItems(lista.SelectedItem.Index).Text
    End If
    Set rs = Nothing
    Set oOC = Nothing
    Exit Sub
fallo:
    MsgBox "Error al cargar los datos del documento. " & Err.Description, vbCritical, App.Title
End Sub

Private Sub cargar_contrato(CONTRATO As Long)
    On Error GoTo fallo
    If CONTRATO > 0 Then
        resizer.Visible = True
        frmContrato.Caption = "Datos del contrato : " & lista.ListItems(lista.SelectedItem.Index).SubItems(2)
    
        Dim oOCD As New clsObras_contratos_detalle
        If oOCD.Carga(CONTRATO) = True Then
            With oOCD
                txtdatos(0) = .getNOMBRE
                txtdatos(1) = .getTELEFONO
                txtdatos(2) = .getOBSERVACIONES
                ' Recepcion
                recepcion_fecha = .getRECEPCION_FECHA
                recepcion_modo.BoundText = .getRECEPCION_MODO
                ' envio
                If .getENVIO_FECHA = "0000-00-00" Then
                    envio_check.Value = Unchecked
                Else
                    envio_check.Value = Checked
                    envio_fecha = .getENVIO_FECHA
                End If
                envio_modo.BoundText = .getENVIO_MODO
                ' Adj
                If .getADJUDICACION_FECHA = "0000-00-00" Then
                    adj_check.Value = Unchecked
                Else
                    adj_check.Value = Checked
                    adj_fecha = .getADJUDICACION_FECHA
                End If
                
                ' Cantidades Contratadas
                cargar_cantidades (CONTRATO)
                ' Ficheros Adjuntos
                cargar_lista_adjuntos (CONTRATO)
                
            End With
        Else
            inicializar_contrato
        End If
    End If
    Set oOCD = Nothing
    Exit Sub
fallo:
    MsgBox "Error al cargar los datos del documento. " & Err.Description, vbCritical, App.Title
End Sub
Private Sub cabecera()
    With lista.ColumnHeaders
        .Add , , "ID", 1, lvwColumnLeft
        .Add , , "Fecha", 1100, lvwColumnLeft
        .Add , , "Descripcion", 6500, lvwColumnLeft
        .Add , , "Estado", 1000, lvwColumnCenter
        .Add , , "Usuario", 1200, lvwColumnLeft
    End With
    With listaAdjuntos.ColumnHeaders
        .Add , , "ID", 1, lvwColumnLeft
        .Add , , "Fecha", 1800, lvwColumnLeft
        .Add , , "Descripcion", 6500, lvwColumnLeft
        .Add , , "Usuario", 1200, lvwColumnLeft
        .Add , , "Documento", 1, lvwColumnLeft
    End With
End Sub

Private Sub lista_Click()
    If lista.ListItems.Count > 0 Then
        cargar_contrato CLng(lista.ListItems(lista.SelectedItem.Index).Text)
    End If
End Sub

Private Sub cargar_combos()
    Dim oDeco As New clsDecodificadora
    oDeco.Cargar_Combo recepcion_modo, DECODIFICADORA.D_CONTRATOS_MODOS_ENTRADA
    oDeco.Cargar_Combo envio_modo, DECODIFICADORA.D_CONTRATOS_MODOS_ENTRADA
End Sub

Private Sub inicializar_contrato()
    txtdatos(0) = ""
    txtdatos(1) = ""
    txtdatos(2) = ""
    
    recepcion_fecha.Value = Date
    recepcion_modo.Text = ""
    
    envio_check.Value = Unchecked
    envio_fecha.Value = Date
    envio_modo.Text = ""
    
    adj_check.Value = Unchecked
    adj_fecha = Date
    
    inicializar_grid
    listaAdjuntos.ListItems.Clear
End Sub

Private Sub cargar_lista_adjuntos(CONTRATO As Long)
    Dim rs_adj As ADODB.Recordset
    Dim oOCA As New clsObras_contratos_adjuntos
    Set rs_adj = oOCA.Listado(CONTRATO)
    listaAdjuntos.ListItems.Clear
    If rs_adj.RecordCount > 0 Then
        Do
            With listaAdjuntos.ListItems.Add(, , rs_adj(0)) ' ID_ADJUNTO
                .SubItems(1) = rs_adj(1) ' FS
                .SubItems(2) = rs_adj(2) ' DESCRIPCION
                .SubItems(3) = rs_adj(3) ' USUARIO
                .SubItems(4) = rs_adj(4) ' DOCUMENTO
            End With
            rs_adj.MoveNext
        Loop Until rs_adj.EOF
    End If
    Set rs_adj = Nothing
    Set oOCA = Nothing
    
End Sub

Private Sub listaAdjuntos_DblClick()
   On Error GoTo listaAdjuntos_DblClick_Error

    If listaAdjuntos.ListItems.Count > 0 Then
        If listaAdjuntos.ListItems(listaAdjuntos.SelectedItem.Index).SubItems(4) <> "" Then
            Dim origen As String
            Dim iret As Long
            origen = ReadINI(App.Path & "\config.ini", "documentos", "ruta") & "\CONTRATOS\" & lista.ListItems(lista.SelectedItem.Index) & "\" & listaAdjuntos.ListItems(listaAdjuntos.SelectedItem.Index).SubItems(4)
            iret = ShellExecute(Me.Hwnd, "Open", origen, "", "", 1)
        End If
    End If

   On Error GoTo 0
   Exit Sub

listaAdjuntos_DblClick_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure listaAdjuntos_DblClick of Formulario frmObras_Contratos"

End Sub
Private Sub cargar_cantidades(CONTRATO As Long)
    Dim oOCC As New clsObras_contratos_cantidades
    Dim rs As ADODB.Recordset
    Set rs = oOCC.Listado(CONTRATO)
    If rs.RecordCount > 0 Then
         Dim fila As Long
         fila = 0
         Do
             xTarifa(fila, ColsTarifa.ID) = CStr(rs(0))
             xTarifa(fila, ColsTarifa.ARTICULO) = CStr(rs(1))
             xTarifa(fila, ColsTarifa.CANTIDAD) = CStr(rs(2))
             xTarifa(fila, ColsTarifa.CANTIDAD_SUM) = CStr(rs(3))
             rs.MoveNext
             fila = fila + 1
         Loop Until rs.EOF
         gridTarifa.Row = 0
         gridTarifa.Col = 0
         gridTarifa.Refresh
     End If
     Set oOCC = Nothing
     Set rs = Nothing
End Sub
Private Sub cargar_articulos()
    Dim rs As ADODB.Recordset
    Dim oArt As New clsArticulos
    Set rs = oArt.ListadoTarifa()
    xarticulos.Clear
    If rs.RecordCount > 0 Then
        xarticulos.ReDim 1, rs.RecordCount, 1, 3
        Dim i As Integer
        i = 1
        Do
            xarticulos(i, 1) = CStr(rs(0))
            xarticulos(i, 2) = CStr(rs(1))
            i = i + 1
            rs.MoveNext
        Loop Until rs.EOF
    Else
        xarticulos.ReDim 1, 1, 1, 3
    End If
    Set tArticulos.Array = xarticulos
    tArticulos.Refresh
    gridTarifa.Refresh
End Sub
Private Sub tArticulos_DropDownClose()
    gridTarifa.Columns(ColsTarifa.ID) = tArticulos.Columns(0)
    gridTarifa.Columns(ColsTarifa.ARTICULO) = tArticulos.Columns(1)
    gridTarifa.Col = 2
End Sub

