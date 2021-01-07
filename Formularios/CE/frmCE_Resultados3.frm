VERSION 5.00
Object = "{0E59F1D2-1FBE-11D0-8FF2-00A0D10038BC}#1.0#0"; "msscript.ocx"
Object = "{A49CE0E0-C0F9-11D2-B0EA-00A024695830}#1.0#0"; "tidate8.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{CEA3B7F6-2847-4E5E-A551-DB7A62489D44}#46.0#0"; "miCombo.ocx"
Object = "{77DDF307-D82B-4757-8B3A-106EC9D75D4B}#8.0#0"; "tdbg8.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#13.2#0"; "Codejock.Controls.v13.2.1.ocx"
Begin VB.Form frmCE_Resultados 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Resultados control de eficacia"
   ClientHeight    =   11265
   ClientLeft      =   1830
   ClientTop       =   1755
   ClientWidth     =   13785
   Icon            =   "frmCE_Resultados2.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   11265
   ScaleWidth      =   13785
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin Geslab.ControlPanelXP cpReactivos 
      Height          =   3975
      Left            =   6930
      TabIndex        =   55
      Top             =   4950
      Width           =   6810
      _ExtentX        =   12012
      _ExtentY        =   7011
      Caption         =   "Reactivos Utilizados en la Muestra"
      BackColor       =   12632256
      TextColor       =   0
      HeaderColor     =   8421504
      PanelOpen       =   0   'False
      Object.Height          =   3975
      Begin VB.Frame frmReactivos 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   0  'None
         Caption         =   "Reactivos"
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
         Height          =   3480
         Left            =   45
         TabIndex        =   56
         Top             =   450
         Width           =   6630
         Begin VB.CommandButton cmdEliminarReactivo 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Eliminar"
            Height          =   795
            Left            =   5670
            Style           =   1  'Graphical
            TabIndex        =   58
            Tag             =   "Elimina el campo seleccionado"
            Top             =   450
            Width           =   915
         End
         Begin VB.CommandButton cmdAnadirReactivo 
            BackColor       =   &H00E0E0E0&
            Caption         =   "A�adir"
            Height          =   750
            Left            =   5670
            Style           =   1  'Graphical
            TabIndex        =   57
            Tag             =   "A�ade campo o modifica el campo existente con el mismo nombre"
            Top             =   1395
            Width           =   915
         End
         Begin MSComctlLib.ListView listaReactivos 
            Height          =   2460
            Left            =   45
            TabIndex        =   59
            Top             =   135
            Width           =   5445
            _ExtentX        =   9604
            _ExtentY        =   4339
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
            TabIndex        =   60
            Top             =   2700
            Width           =   5280
            _ExtentX        =   9313
            _ExtentY        =   582
         End
         Begin pryCombo.miCombo cmbReactivosInternos 
            Height          =   330
            Left            =   765
            TabIndex        =   62
            Top             =   3060
            Width           =   5280
            _ExtentX        =   9313
            _ExtentY        =   582
         End
         Begin VB.Label lblCampos 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C0C0&
            Caption         =   "Interno"
            Height          =   195
            Index           =   8
            Left            =   90
            TabIndex        =   63
            Top             =   3105
            Width           =   495
         End
         Begin VB.Label lblCampos 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C0C0&
            Caption         =   "Externo"
            Height          =   195
            Index           =   5
            Left            =   90
            TabIndex        =   61
            Top             =   2745
            Width           =   540
         End
      End
   End
   Begin Geslab.ControlPanelXP ControlPanelXP2 
      Height          =   4425
      Left            =   45
      TabIndex        =   64
      Top             =   4950
      Width           =   6810
      _ExtentX        =   12012
      _ExtentY        =   7805
      Caption         =   "Equipos Utilizados en la Muestra"
      BackColor       =   12632256
      TextColor       =   0
      HeaderColor     =   8421504
      PanelOpen       =   0   'False
      Object.Height          =   4425
      Begin VB.Frame frmEquipos 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
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
         ForeColor       =   &H80000008&
         Height          =   3930
         Left            =   90
         TabIndex        =   65
         Top             =   405
         Width           =   6585
         Begin VB.CommandButton cmdModificarEquipo 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Modificar"
            Height          =   765
            Left            =   4590
            Style           =   1  'Graphical
            TabIndex        =   96
            Tag             =   "A�ade campo o modifica el campo existente con el mismo nombre"
            Top             =   3105
            Width           =   975
         End
         Begin VB.TextBox txtusos 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            Height          =   285
            Left            =   5715
            TabIndex        =   94
            Top             =   2700
            Width           =   780
         End
         Begin VB.CommandButton cmdVerificacion 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Verificaci�n"
            Height          =   765
            Left            =   5580
            Style           =   1  'Graphical
            TabIndex        =   71
            Tag             =   "A�ade campo o modifica el campo existente con el mismo nombre"
            Top             =   3105
            Width           =   975
         End
         Begin VB.CommandButton cmdEliminarEquipo 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Eliminar"
            Height          =   810
            Left            =   0
            Style           =   1  'Graphical
            TabIndex        =   67
            Tag             =   "Elimina el campo seleccionado"
            Top             =   3060
            Width           =   975
         End
         Begin VB.CommandButton cmdAnadirEquipo 
            BackColor       =   &H00E0E0E0&
            Caption         =   "A�adir"
            Height          =   765
            Left            =   3600
            Style           =   1  'Graphical
            TabIndex        =   66
            Tag             =   "A�ade campo o modifica el campo existente con el mismo nombre"
            Top             =   3105
            Width           =   975
         End
         Begin MSComctlLib.ListView listaEquipos 
            Height          =   2355
            Left            =   0
            TabIndex        =   68
            Top             =   270
            Width           =   6510
            _ExtentX        =   11483
            _ExtentY        =   4154
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
            TabIndex        =   69
            Top             =   2700
            Width           =   5010
            _ExtentX        =   8837
            _ExtentY        =   582
         End
         Begin VB.Label Label1 
            BackColor       =   &H00C0C0C0&
            Caption         =   "Usos"
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
            Index           =   6
            Left            =   5175
            TabIndex        =   95
            Top             =   2745
            Width           =   810
         End
         Begin VB.Label Label2 
            BackColor       =   &H00C0C0C0&
            Caption         =   "Marque los equipos que deben salir en el informe"
            ForeColor       =   &H000000FF&
            Height          =   195
            Left            =   0
            TabIndex        =   70
            Top             =   45
            Width           =   4335
         End
      End
   End
   Begin VB.Frame frmResultados 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   "Resultados"
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
      Height          =   4560
      Left            =   45
      TabIndex        =   36
      Top             =   5445
      Width           =   13695
      Begin MSComctlLib.ListView auxdatos 
         Height          =   1710
         Left            =   6885
         TabIndex        =   37
         Top             =   2475
         Visible         =   0   'False
         Width           =   6315
         _ExtentX        =   11139
         _ExtentY        =   3016
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         AllowReorder    =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   12640511
         BorderStyle     =   1
         Appearance      =   0
         NumItems        =   0
      End
      Begin MSComctlLib.ListView lista 
         Height          =   3960
         Left            =   45
         TabIndex        =   43
         Top             =   540
         Width           =   6990
         _ExtentX        =   12330
         _ExtentY        =   6985
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         AllowReorder    =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   13230796
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
      Begin VB.Frame Frame3 
         BackColor       =   &H00C0C0C0&
         Height          =   720
         Left            =   7065
         TabIndex        =   38
         Top             =   3825
         Width           =   6495
         Begin VB.CommandButton cmdcalcular 
            BackColor       =   &H00E0E0E0&
            Enabled         =   0   'False
            Height          =   555
            Left            =   5895
            Style           =   1  'Graphical
            TabIndex        =   47
            Top             =   135
            Width           =   555
         End
         Begin VB.TextBox txtvalor 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            Height          =   330
            Left            =   4095
            TabIndex        =   40
            Top             =   225
            Width           =   1635
         End
         Begin VB.TextBox txtdato 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            Height          =   330
            Left            =   765
            Locked          =   -1  'True
            TabIndex        =   39
            Top             =   225
            Width           =   2715
         End
         Begin VB.Label Label1 
            BackColor       =   &H00C0C0C0&
            Caption         =   "Valor"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   255
            Index           =   3
            Left            =   3555
            TabIndex        =   42
            Top             =   270
            Width           =   555
         End
         Begin VB.Label Label1 
            BackColor       =   &H00C0C0C0&
            Caption         =   "Campo"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   255
            Index           =   0
            Left            =   90
            TabIndex        =   41
            Top             =   315
            Width           =   855
         End
      End
      Begin MSComctlLib.ListView datos 
         Height          =   3195
         Left            =   7065
         TabIndex        =   44
         Top             =   540
         Width           =   6480
         _ExtentX        =   11430
         _ExtentY        =   5636
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
      Begin VB.Label lblmsg 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         Caption         =   "Probetas"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   300
         Left            =   45
         TabIndex        =   46
         Top             =   225
         Width           =   7035
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Campos"
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
         Height          =   285
         Index           =   0
         Left            =   7065
         TabIndex        =   45
         Top             =   225
         Width           =   6465
      End
      Begin VB.Label lblestado 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0FF&
         Caption         =   "DUPLICADA"
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
         Height          =   285
         Left            =   10845
         TabIndex        =   48
         Top             =   225
         Width           =   2715
      End
   End
   Begin TrueDBGrid80.TDBGrid gridP 
      Height          =   4305
      Left            =   45
      TabIndex        =   0
      Top             =   5535
      Width           =   13650
      _ExtentX        =   24077
      _ExtentY        =   7594
      _LayoutType     =   4
      _RowHeight      =   -2147483647
      _WasPersistedAsPixels=   0
      Columns(0)._VlistStyle=   0
      Columns(0)._MaxComboItems=   5
      Columns(0).Caption=   "Identificaci�n Canagrosa"
      Columns(0).DataField=   ""
      Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(1)._VlistStyle=   0
      Columns(1)._MaxComboItems=   5
      Columns(1).Caption=   "Identificaci�n Cliente"
      Columns(1).DataField=   ""
      Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(2)._VlistStyle=   0
      Columns(2)._MaxComboItems=   5
      Columns(2).Caption=   "Dimensi�n"
      Columns(2).DataField=   ""
      Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(3)._VlistStyle=   0
      Columns(3)._MaxComboItems=   5
      Columns(3).Caption=   "Fecha"
      Columns(3).DataField=   ""
      Columns(3).NumberFormat=   "General Date"
      Columns(3).ExternalEditor=   "TDBDate1"
      Columns(3).ExternalEditor.vt=   8
      Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(4)._VlistStyle=   0
      Columns(4)._MaxComboItems=   5
      Columns(4).Caption=   "Resultado"
      Columns(4).DataField=   ""
      Columns(4).DropDown=   "tResponsables"
      Columns(4).DropDown.vt=   8
      Columns(4)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(5)._VlistStyle=   2
      Columns(5)._MaxComboItems=   5
      Columns(5).ValueItems(0)._DefaultItem=   0
      Columns(5).ValueItems(0).Value=   "Si"
      Columns(5).ValueItems(0).Value.vt=   8
      Columns(5).ValueItems(0).DisplayValue=   "SI"
      Columns(5).ValueItems(0).DisplayValue.vt=   8
      Columns(5).ValueItems(0)._PropDict=   "_DefaultItem,517,2"
      Columns(5).ValueItems(1)._DefaultItem=   0
      Columns(5).ValueItems(1).Value=   "No"
      Columns(5).ValueItems(1).Value.vt=   8
      Columns(5).ValueItems(1).DisplayValue=   "NO"
      Columns(5).ValueItems(1).DisplayValue.vt=   8
      Columns(5).ValueItems(1)._PropDict=   "_DefaultItem,517,2"
      Columns(5).ValueItems.Count=   2
      Columns(5).Caption=   "Conforme"
      Columns(5).DataField=   ""
      Columns(5)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(6)._VlistStyle=   0
      Columns(6)._MaxComboItems=   5
      Columns(6).Caption=   "Designacion"
      Columns(6).DataField=   ""
      Columns(6).ExternalEditor=   "TDBDate1"
      Columns(6).ExternalEditor.vt=   8
      Columns(6)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(7)._VlistStyle=   0
      Columns(7)._MaxComboItems=   5
      Columns(7).Caption=   "Probeta"
      Columns(7).DataField=   ""
      Columns(7)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(8)._VlistStyle=   0
      Columns(8)._MaxComboItems=   5
      Columns(8).Caption=   "Area"
      Columns(8).DataField=   ""
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
      Splits(0)._ColumnProps(1)=   "Column(0).Width=5689"
      Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
      Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=5609"
      Splits(0)._ColumnProps(4)=   "Column(0)._EditAlways=0"
      Splits(0)._ColumnProps(5)=   "Column(0)._ColStyle=8192"
      Splits(0)._ColumnProps(6)=   "Column(0).AllowFocus=0"
      Splits(0)._ColumnProps(7)=   "Column(0).Order=1"
      Splits(0)._ColumnProps(8)=   "Column(1).Width=5503"
      Splits(0)._ColumnProps(9)=   "Column(1).DividerColor=0"
      Splits(0)._ColumnProps(10)=   "Column(1)._WidthInPix=5424"
      Splits(0)._ColumnProps(11)=   "Column(1)._EditAlways=0"
      Splits(0)._ColumnProps(12)=   "Column(1)._ColStyle=8196"
      Splits(0)._ColumnProps(13)=   "Column(1).Order=2"
      Splits(0)._ColumnProps(14)=   "Column(2).Width=3995"
      Splits(0)._ColumnProps(15)=   "Column(2).DividerColor=0"
      Splits(0)._ColumnProps(16)=   "Column(2)._WidthInPix=3916"
      Splits(0)._ColumnProps(17)=   "Column(2)._EditAlways=0"
      Splits(0)._ColumnProps(18)=   "Column(2)._ColStyle=8196"
      Splits(0)._ColumnProps(19)=   "Column(2).Order=3"
      Splits(0)._ColumnProps(20)=   "Column(3).Width=2302"
      Splits(0)._ColumnProps(21)=   "Column(3).DividerColor=0"
      Splits(0)._ColumnProps(22)=   "Column(3)._WidthInPix=2223"
      Splits(0)._ColumnProps(23)=   "Column(3)._EditAlways=0"
      Splits(0)._ColumnProps(24)=   "Column(3)._ColStyle=1"
      Splits(0)._ColumnProps(25)=   "Column(3).Order=4"
      Splits(0)._ColumnProps(26)=   "Column(4).Width=3916"
      Splits(0)._ColumnProps(27)=   "Column(4).DividerColor=0"
      Splits(0)._ColumnProps(28)=   "Column(4)._WidthInPix=3836"
      Splits(0)._ColumnProps(29)=   "Column(4)._EditAlways=0"
      Splits(0)._ColumnProps(30)=   "Column(4)._ColStyle=1"
      Splits(0)._ColumnProps(31)=   "Column(4).Order=5"
      Splits(0)._ColumnProps(32)=   "Column(4).DropDownList=1"
      Splits(0)._ColumnProps(33)=   "Column(5).Width=159"
      Splits(0)._ColumnProps(34)=   "Column(5).DividerColor=0"
      Splits(0)._ColumnProps(35)=   "Column(5)._WidthInPix=79"
      Splits(0)._ColumnProps(36)=   "Column(5)._EditAlways=0"
      Splits(0)._ColumnProps(37)=   "Column(5)._ColStyle=1"
      Splits(0)._ColumnProps(38)=   "Column(5).Order=6"
      Splits(0)._ColumnProps(39)=   "Column(6).Width=1402"
      Splits(0)._ColumnProps(40)=   "Column(6).DividerColor=0"
      Splits(0)._ColumnProps(41)=   "Column(6)._WidthInPix=1323"
      Splits(0)._ColumnProps(42)=   "Column(6)._EditAlways=0"
      Splits(0)._ColumnProps(43)=   "Column(6)._ColStyle=8193"
      Splits(0)._ColumnProps(44)=   "Column(6).Visible=0"
      Splits(0)._ColumnProps(45)=   "Column(6).Order=7"
      Splits(0)._ColumnProps(46)=   "Column(7).Width=1640"
      Splits(0)._ColumnProps(47)=   "Column(7).DividerColor=0"
      Splits(0)._ColumnProps(48)=   "Column(7)._WidthInPix=1561"
      Splits(0)._ColumnProps(49)=   "Column(7)._EditAlways=0"
      Splits(0)._ColumnProps(50)=   "Column(7)._ColStyle=1"
      Splits(0)._ColumnProps(51)=   "Column(7).Visible=0"
      Splits(0)._ColumnProps(52)=   "Column(7).Order=8"
      Splits(0)._ColumnProps(53)=   "Column(8).Width=185"
      Splits(0)._ColumnProps(54)=   "Column(8).DividerColor=0"
      Splits(0)._ColumnProps(55)=   "Column(8)._WidthInPix=106"
      Splits(0)._ColumnProps(56)=   "Column(8)._EditAlways=0"
      Splits(0)._ColumnProps(57)=   "Column(8)._ColStyle=1"
      Splits(0)._ColumnProps(58)=   "Column(8).Visible=0"
      Splits(0)._ColumnProps(59)=   "Column(8).Order=9"
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
      Caption         =   "Resultados Probetas"
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
      _StyleDefs(36)  =   "Splits(0).Columns(0).Style:id=24,.parent=11,.alignment=0,.fgcolor=&HFF&"
      _StyleDefs(37)  =   ":id=24,.locked=-1,.bold=-1,.fontsize=975,.italic=0,.underline=0,.strikethrough=0"
      _StyleDefs(38)  =   ":id=24,.charset=0"
      _StyleDefs(39)  =   ":id=24,.fontname=MS Sans Serif"
      _StyleDefs(40)  =   "Splits(0).Columns(0).HeadingStyle:id=21,.parent=12"
      _StyleDefs(41)  =   "Splits(0).Columns(0).FooterStyle:id=22,.parent=13"
      _StyleDefs(42)  =   "Splits(0).Columns(0).EditorStyle:id=23,.parent=15,.bold=0,.fontsize=975"
      _StyleDefs(43)  =   ":id=23,.italic=0,.underline=0,.strikethrough=0,.charset=0"
      _StyleDefs(44)  =   ":id=23,.fontname=MS Sans Serif"
      _StyleDefs(45)  =   "Splits(0).Columns(1).Style:id=66,.parent=11,.locked=-1"
      _StyleDefs(46)  =   "Splits(0).Columns(1).HeadingStyle:id=63,.parent=12"
      _StyleDefs(47)  =   "Splits(0).Columns(1).FooterStyle:id=64,.parent=13"
      _StyleDefs(48)  =   "Splits(0).Columns(1).EditorStyle:id=65,.parent=15"
      _StyleDefs(49)  =   "Splits(0).Columns(2).Style:id=70,.parent=11,.locked=-1"
      _StyleDefs(50)  =   "Splits(0).Columns(2).HeadingStyle:id=67,.parent=12"
      _StyleDefs(51)  =   "Splits(0).Columns(2).FooterStyle:id=68,.parent=13"
      _StyleDefs(52)  =   "Splits(0).Columns(2).EditorStyle:id=69,.parent=15"
      _StyleDefs(53)  =   "Splits(0).Columns(3).Style:id=32,.parent=11,.alignment=2"
      _StyleDefs(54)  =   "Splits(0).Columns(3).HeadingStyle:id=29,.parent=12"
      _StyleDefs(55)  =   "Splits(0).Columns(3).FooterStyle:id=30,.parent=13"
      _StyleDefs(56)  =   "Splits(0).Columns(3).EditorStyle:id=31,.parent=15"
      _StyleDefs(57)  =   "Splits(0).Columns(4).Style:id=36,.parent=11,.alignment=2,.bgcolor=&HC1FFFF&"
      _StyleDefs(58)  =   "Splits(0).Columns(4).HeadingStyle:id=33,.parent=12"
      _StyleDefs(59)  =   "Splits(0).Columns(4).FooterStyle:id=34,.parent=13"
      _StyleDefs(60)  =   "Splits(0).Columns(4).EditorStyle:id=35,.parent=15"
      _StyleDefs(61)  =   "Splits(0).Columns(5).Style:id=54,.parent=11,.alignment=2,.bgcolor=&HC1FFFF&"
      _StyleDefs(62)  =   "Splits(0).Columns(5).HeadingStyle:id=51,.parent=12"
      _StyleDefs(63)  =   "Splits(0).Columns(5).FooterStyle:id=52,.parent=13"
      _StyleDefs(64)  =   "Splits(0).Columns(5).EditorStyle:id=53,.parent=15"
      _StyleDefs(65)  =   "Splits(0).Columns(6).Style:id=28,.parent=11,.alignment=2,.locked=-1"
      _StyleDefs(66)  =   "Splits(0).Columns(6).HeadingStyle:id=25,.parent=12"
      _StyleDefs(67)  =   "Splits(0).Columns(6).FooterStyle:id=26,.parent=13"
      _StyleDefs(68)  =   "Splits(0).Columns(6).EditorStyle:id=27,.parent=15"
      _StyleDefs(69)  =   "Splits(0).Columns(7).Style:id=58,.parent=11,.alignment=2"
      _StyleDefs(70)  =   "Splits(0).Columns(7).HeadingStyle:id=55,.parent=12"
      _StyleDefs(71)  =   "Splits(0).Columns(7).FooterStyle:id=56,.parent=13"
      _StyleDefs(72)  =   "Splits(0).Columns(7).EditorStyle:id=57,.parent=15"
      _StyleDefs(73)  =   "Splits(0).Columns(8).Style:id=62,.parent=11,.alignment=2,.bgcolor=&H80FFFF&"
      _StyleDefs(74)  =   "Splits(0).Columns(8).HeadingStyle:id=59,.parent=12"
      _StyleDefs(75)  =   "Splits(0).Columns(8).FooterStyle:id=60,.parent=13"
      _StyleDefs(76)  =   "Splits(0).Columns(8).EditorStyle:id=61,.parent=15"
      _StyleDefs(77)  =   "Named:id=37:Normal"
      _StyleDefs(78)  =   ":id=37,.parent=0,.alignment=3,.bgcolor=&H80000018&,.bgpicMode=2,.bold=0"
      _StyleDefs(79)  =   ":id=37,.fontsize=825,.italic=0,.underline=0,.strikethrough=0,.charset=0"
      _StyleDefs(80)  =   ":id=37,.fontname=MS Sans Serif"
      _StyleDefs(81)  =   "Named:id=38:Heading"
      _StyleDefs(82)  =   ":id=38,.parent=37,.valignment=2,.bgcolor=&H80000016&,.fgcolor=&H80000012&"
      _StyleDefs(83)  =   ":id=38,.wraptext=-1,.bold=0,.fontsize=825,.italic=0,.underline=0"
      _StyleDefs(84)  =   ":id=38,.strikethrough=0,.charset=0"
      _StyleDefs(85)  =   ":id=38,.fontname=MS Sans Serif"
      _StyleDefs(86)  =   "Named:id=39:Footing"
      _StyleDefs(87)  =   ":id=39,.parent=37,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(88)  =   "Named:id=40:Selected"
      _StyleDefs(89)  =   ":id=40,.parent=37,.bgcolor=&H8080FF&,.fgcolor=&H0&,.bold=0,.fontsize=825"
      _StyleDefs(90)  =   ":id=40,.italic=0,.underline=0,.strikethrough=0,.charset=0"
      _StyleDefs(91)  =   ":id=40,.fontname=MS Sans Serif"
      _StyleDefs(92)  =   "Named:id=41:Caption"
      _StyleDefs(93)  =   ":id=41,.parent=38,.alignment=2"
      _StyleDefs(94)  =   "Named:id=42:HighlightRow"
      _StyleDefs(95)  =   ":id=42,.parent=37,.bgcolor=&H80000008&,.fgcolor=&H80000005&"
      _StyleDefs(96)  =   "Named:id=43:EvenRow"
      _StyleDefs(97)  =   ":id=43,.parent=37,.bgcolor=&HFFFFFF&,.wraptext=-1"
      _StyleDefs(98)  =   "Named:id=44:OddRow"
      _StyleDefs(99)  =   ":id=44,.parent=37,.bgcolor=&HD5ECF9&"
      _StyleDefs(100) =   "Named:id=47:RecordSelector"
      _StyleDefs(101) =   ":id=47,.parent=38"
      _StyleDefs(102) =   "Named:id=50:FilterBar"
      _StyleDefs(103) =   ":id=50,.parent=37"
   End
   Begin VB.CommandButton cmdCurvas 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Hist�rico"
      Height          =   825
      Left            =   3420
      Style           =   1  'Graphical
      TabIndex        =   77
      Top             =   10395
      Width           =   1095
   End
   Begin VB.CheckBox chkModificar 
      Caption         =   "Permiso Modificar Cerrada"
      Height          =   195
      Left            =   9315
      TabIndex        =   76
      Top             =   10350
      Visible         =   0   'False
      Width           =   1230
   End
   Begin VB.Frame frmSPDA 
      BackColor       =   &H00C0C0C0&
      Caption         =   "SPDA"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   825
      Left            =   4815
      TabIndex        =   72
      Top             =   10395
      Width           =   4695
      Begin VB.TextBox txtSPDA 
         Alignment       =   2  'Center
         Height          =   330
         Left            =   1575
         TabIndex        =   73
         Top             =   315
         Width           =   1635
      End
      Begin XtremeSuiteControls.PushButton cmdSpdaReiniciar 
         Height          =   300
         Left            =   3330
         TabIndex        =   75
         Top             =   315
         Width           =   1140
         _Version        =   851970
         _ExtentX        =   2011
         _ExtentY        =   529
         _StockProps     =   79
         Caption         =   "Reiniciar"
         Appearance      =   5
         Picture         =   "frmCE_Resultados2.frx":08CA
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Acumulado (mg/dl)"
         Height          =   195
         Index           =   9
         Left            =   135
         TabIndex        =   74
         Top             =   360
         Width           =   1335
      End
   End
   Begin Geslab.ControlPanelXP ControlPanelXP1 
      Height          =   2040
      Left            =   45
      TabIndex        =   18
      Top             =   585
      Width           =   13695
      _ExtentX        =   24156
      _ExtentY        =   3598
      Caption         =   "Datos definidos del Tipo de Ensayo de Eficacia"
      BackColor       =   12632256
      TextColor       =   0
      HeaderColor     =   8421504
      Object.Height          =   2040
      Begin VB.Frame Frame2 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   0  'None
         Caption         =   "Datos definidos del Tipo de Ensayo de Eficacia"
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
         Height          =   1590
         Left            =   135
         TabIndex        =   19
         Top             =   405
         Width           =   13245
         Begin VB.CheckBox chkLote 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            Caption         =   "Lote Probetas"
            ForeColor       =   &H80000008&
            Height          =   240
            Left            =   11430
            TabIndex        =   27
            Top             =   540
            Width           =   1365
         End
         Begin VB.CheckBox chkEspesor 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            Caption         =   "Incluye Espesor"
            ForeColor       =   &H80000008&
            Height          =   240
            Left            =   11430
            TabIndex        =   26
            Top             =   855
            Width           =   1500
         End
         Begin VB.TextBox txtDatos 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            Height          =   285
            Index           =   4
            Left            =   12105
            Locked          =   -1  'True
            TabIndex        =   25
            Top             =   135
            Width           =   1020
         End
         Begin VB.TextBox txtDatos 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            Height          =   285
            Index           =   3
            Left            =   10170
            Locked          =   -1  'True
            TabIndex        =   24
            Top             =   135
            Width           =   975
         End
         Begin VB.TextBox txthoras 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            Height          =   285
            Left            =   10170
            Locked          =   -1  'True
            TabIndex        =   23
            Top             =   810
            Width           =   990
         End
         Begin VB.TextBox txtDatos 
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            Height          =   285
            Index           =   2
            Left            =   1485
            Locked          =   -1  'True
            TabIndex        =   22
            Top             =   945
            Width           =   7590
         End
         Begin VB.TextBox txtDatos 
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            Height          =   510
            Index           =   1
            Left            =   1485
            Locked          =   -1  'True
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   21
            Top             =   405
            Width           =   7590
         End
         Begin VB.TextBox txtDatos 
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            Height          =   285
            Index           =   5
            Left            =   1485
            Locked          =   -1  'True
            TabIndex        =   20
            Top             =   1260
            Width           =   7590
         End
         Begin pryCombo.miCombo cmbTipoEnsayo 
            Height          =   330
            Left            =   1485
            TabIndex        =   78
            Top             =   45
            Width           =   7575
            _ExtentX        =   13361
            _ExtentY        =   582
         End
         Begin XtremeSuiteControls.PushButton cmdModificarEnsayo 
            Height          =   300
            Left            =   9675
            TabIndex        =   79
            Top             =   1260
            Width           =   2940
            _Version        =   851970
            _ExtentX        =   5186
            _ExtentY        =   529
            _StockProps     =   79
            Caption         =   "Modificar Tipo de Ensayo"
            Appearance      =   5
            Picture         =   "frmCE_Resultados2.frx":712C
         End
         Begin VB.Label lblCampos 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C0C0&
            Caption         =   "Aceptaci�n (C.A)"
            Height          =   195
            Index           =   16
            Left            =   90
            TabIndex        =   35
            Top             =   630
            Width           =   1200
         End
         Begin VB.Label lblCampos 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C0C0&
            Caption         =   "Rango Max."
            Height          =   195
            Index           =   15
            Left            =   11205
            TabIndex        =   34
            Top             =   180
            Width           =   870
         End
         Begin VB.Label lblCampos 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C0C0&
            Caption         =   "Rango Min."
            Height          =   195
            Index           =   14
            Left            =   9225
            TabIndex        =   33
            Top             =   180
            Width           =   825
         End
         Begin VB.Label lblCampos 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C0C0&
            Caption         =   "Duraci�n (h)"
            Height          =   195
            Index           =   13
            Left            =   9225
            TabIndex        =   32
            Top             =   855
            Width           =   870
         End
         Begin VB.Label lblCampos 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C0C0&
            Caption         =   "Norma (C.A)"
            Height          =   195
            Index           =   7
            Left            =   90
            TabIndex        =   31
            Top             =   990
            Width           =   855
         End
         Begin VB.Label lblCampos 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C0C0&
            Caption         =   "Criterio"
            Height          =   195
            Index           =   0
            Left            =   90
            TabIndex        =   30
            Top             =   405
            Width           =   480
         End
         Begin VB.Label lblCampos 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C0C0&
            Caption         =   "Nombre"
            Height          =   195
            Index           =   1
            Left            =   90
            TabIndex        =   29
            Top             =   90
            Width           =   555
         End
         Begin VB.Label lblCampos 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C0C0&
            Caption         =   "P.N.T."
            Height          =   195
            Index           =   3
            Left            =   90
            TabIndex        =   28
            Top             =   1305
            Width           =   465
         End
      End
   End
   Begin VB.CheckBox chkDuplicada 
      Caption         =   "Duplicada"
      Height          =   195
      Left            =   8190
      TabIndex        =   50
      Top             =   10440
      Visible         =   0   'False
      Width           =   1050
   End
   Begin VB.TextBox txtformula 
      Height          =   285
      Left            =   7335
      TabIndex        =   49
      Text            =   "0"
      Top             =   10305
      Visible         =   0   'False
      Width           =   510
   End
   Begin VB.CommandButton cmdObservador 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Observador"
      Height          =   825
      Left            =   10365
      Style           =   1  'Graphical
      TabIndex        =   16
      Tag             =   "A�ade campo o modifica el campo existente con el mismo nombre"
      Top             =   10395
      Width           =   1095
   End
   Begin VB.CommandButton cmdtipoensayo 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Tipo Ensayo"
      Height          =   825
      Left            =   2295
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   10395
      Width           =   1095
   End
   Begin VB.TextBox txttipoensayo 
      Height          =   285
      Left            =   6795
      TabIndex        =   14
      Text            =   "0"
      Top             =   10305
      Visible         =   0   'False
      Width           =   510
   End
   Begin VB.CommandButton cmdPNT 
      BackColor       =   &H00E0E0E0&
      Caption         =   "P.N.T."
      Height          =   825
      Left            =   1170
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   10395
      Width           =   1095
   End
   Begin VB.CommandButton cmdImagen 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Imagenes"
      Height          =   825
      Left            =   45
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   10395
      Width           =   1095
   End
   Begin VB.TextBox txtnumprobetas 
      Height          =   285
      Left            =   6255
      TabIndex        =   10
      Text            =   "Text1"
      Top             =   10305
      Visible         =   0   'False
      Width           =   510
   End
   Begin VB.CommandButton cmdAceptar 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Aceptar"
      Height          =   825
      Left            =   11520
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   10395
      Width           =   1095
   End
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      ForeColor       =   &H80000008&
      Height          =   2280
      Left            =   45
      TabIndex        =   2
      Top             =   2655
      Width           =   13695
      Begin VB.Frame Frame4 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Caption         =   "Tiempo del ensayo"
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
         Left            =   3735
         TabIndex        =   86
         Top             =   1530
         Width           =   9915
         Begin VB.CommandButton cmdComienzo 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Comenzar"
            Height          =   420
            Left            =   8280
            Style           =   1  'Graphical
            TabIndex        =   87
            Top             =   180
            Width           =   1590
         End
         Begin MSComCtl2.DTPicker ddesde 
            Height          =   330
            Left            =   1305
            TabIndex        =   88
            Top             =   225
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
            Format          =   61014017
            CurrentDate     =   38000
            MinDate         =   2
         End
         Begin MSComCtl2.DTPicker dhdesde 
            Height          =   330
            Left            =   2655
            TabIndex        =   89
            Top             =   225
            Width           =   1155
            _ExtentX        =   2037
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
            CustomFormat    =   "00:00:00"
            Format          =   61014018
            CurrentDate     =   38000
            MinDate         =   2
         End
         Begin MSComCtl2.DTPicker dhasta 
            Height          =   330
            Left            =   5085
            TabIndex        =   90
            Top             =   225
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
            Format          =   61014017
            CurrentDate     =   38000
            MinDate         =   2
         End
         Begin MSComCtl2.DTPicker dhhasta 
            Height          =   330
            Left            =   6480
            TabIndex        =   91
            Top             =   225
            Width           =   1155
            _ExtentX        =   2037
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
            Format          =   61014018
            CurrentDate     =   38000
            MinDate         =   2
         End
         Begin VB.Label lblCampos 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C0C0&
            Caption         =   "Fecha de fin"
            Height          =   195
            Index           =   2
            Left            =   4095
            TabIndex        =   93
            Top             =   315
            Width           =   885
         End
         Begin VB.Label lblCampos 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C0C0&
            Caption         =   "Fecha de inicio"
            Height          =   195
            Index           =   4
            Left            =   90
            TabIndex        =   92
            Top             =   315
            Width           =   1080
         End
      End
      Begin VB.Frame frmFechasEnsayo 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Caption         =   "Fechas del ensayo"
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
         Left            =   90
         TabIndex        =   82
         Top             =   1530
         Width           =   3750
         Begin VB.CheckBox chkSinEspecificar 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            Caption         =   "Sin Especificar"
            ForeColor       =   &H80000008&
            Height          =   330
            Left            =   2250
            TabIndex        =   83
            Top             =   270
            Width           =   1365
         End
         Begin MSComCtl2.DTPicker fprocesado 
            Height          =   330
            Left            =   900
            TabIndex        =   84
            Top             =   270
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
            Format          =   61014017
            CurrentDate     =   38000
            MinDate         =   2
         End
         Begin VB.Label lblCampos 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C0C0&
            Caption         =   "Procesado"
            Height          =   195
            Index           =   6
            Left            =   45
            TabIndex        =   85
            Top             =   315
            Width           =   765
         End
      End
      Begin VB.TextBox txtMaterial 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   1260
         TabIndex        =   54
         Top             =   855
         Width           =   9645
      End
      Begin VB.TextBox txtCondicionesAmbientales 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   8820
         TabIndex        =   51
         Top             =   180
         Width           =   2130
      End
      Begin VB.CommandButton cmdmodificarprobetas 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Datos Probetas"
         Height          =   735
         Left            =   12015
         Picture         =   "frmCE_Resultados2.frx":D98E
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   450
         Width           =   1605
      End
      Begin VB.TextBox txtespesor 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         Height          =   285
         Left            =   1260
         TabIndex        =   3
         Top             =   180
         Width           =   5235
      End
      Begin pryCombo.miCombo cmbLote 
         Height          =   375
         Left            =   1260
         TabIndex        =   53
         Top             =   495
         Width           =   10725
         _ExtentX        =   18918
         _ExtentY        =   661
      End
      Begin pryCombo.miCombo cmbDimension 
         Height          =   375
         Left            =   1260
         TabIndex        =   97
         Top             =   1170
         Width           =   10725
         _ExtentX        =   18918
         _ExtentY        =   661
      End
      Begin VB.Label Label1 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Dimensiones"
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
         Left            =   90
         TabIndex        =   81
         Top             =   1215
         Width           =   1215
      End
      Begin VB.Label Label1 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Material"
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
         Index           =   2
         Left            =   90
         TabIndex        =   80
         Top             =   900
         Width           =   1035
      End
      Begin VB.Label Label1 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Condiciones Ambientales"
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
         Left            =   6570
         TabIndex        =   52
         Top             =   225
         Width           =   2340
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Lote Probetas"
         Height          =   195
         Index           =   18
         Left            =   90
         TabIndex        =   9
         Top             =   540
         Width           =   990
      End
      Begin VB.Label Label1 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Espesor"
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
         Index           =   5
         Left            =   90
         TabIndex        =   4
         Top             =   225
         Width           =   1170
      End
   End
   Begin VB.CommandButton cmdSalir 
      BackColor       =   &H00E0E0E0&
      Caption         =   "ESC-Salir"
      Height          =   825
      Left            =   12645
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   10395
      Width           =   1095
   End
   Begin TDBDate6Ctl.TDBDate TDBDate1 
      Height          =   255
      Left            =   7785
      TabIndex        =   7
      Top             =   10305
      Visible         =   0   'False
      Width           =   1185
      _Version        =   65536
      _ExtentX        =   2090
      _ExtentY        =   450
      Calendar        =   "frmCE_Resultados2.frx":E1DC
      Caption         =   "frmCE_Resultados2.frx":E2F4
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmCE_Resultados2.frx":E360
      Keys            =   "frmCE_Resultados2.frx":E37E
      Spin            =   "frmCE_Resultados2.frx":E3DC
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
   Begin MSScriptControlCtl.ScriptControl sc 
      Left            =   4410
      Top             =   10665
      _ExtentX        =   1005
      _ExtentY        =   1005
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
      Left            =   11565
      TabIndex        =   17
      Top             =   90
      Width           =   2085
   End
   Begin VB.Label lbltitulo 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "Resultados de Control de Eficacia"
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
      TabIndex        =   8
      Top             =   120
      Width           =   3555
   End
   Begin VB.Label lblmensaje 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   "Probetas"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   495
      TabIndex        =   6
      Top             =   10080
      Width           =   12840
   End
   Begin VB.Shape fondo 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   555
      Left            =   0
      Top             =   0
      Width           =   13770
   End
End
Attribute VB_Name = "frmCE_Resultados"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Private WithEvents TecladoNumerico As frmTecladoNumerico
'Private blnEsTablet As Boolean
Private mvarlngIdTipoEnsayo As Long

Public PK_ID_MUESTRA As Long
Dim xP As New XArrayDB
Const filasP As Integer = 100
Const ColP As Integer = 8

'Private blnPrimeraVez As Boolean
Private Enum ColsP
    Identificacion = 0
    IDENTIFICACION_CLIENTE = 1
    DIMENSION = 2
    fecha = 3
    RESULTADO = 4
    CONFORME = 5
    DESIGNACION = 6
    PROBETA = 7
    AREA = 8
End Enum

'Public EQUIPOS_MODIFICADOS As Boolean

Private mvarblnMuestra_Cerrada  As Boolean
Private Sub cmdCurvas_Click()
    frmHistoricoDeterminacionCE.ID_MUESTRA = PK_ID_MUESTRA
    frmHistoricoDeterminacionCE.Show 1
End Sub
Private Sub cmdCalcular_Click()
    On Error GoTo fallo
    Dim i As Integer
    Dim cont As Integer
    cont = 0
    Dim requeridos As Boolean
    requeridos = True
    ' Validamos los campos requeridos para el calculo
    For i = datos.selectedItem.Index To 1 Step -1
         If datos.ListItems(i).bold = False Then
             If Trim(datos.ListItems(i).SubItems(1)) = "" Then
                 requeridos = False
             End If
         End If
    Next
    ' Comprobamos que esten todos los campos requeridos
    If requeridos = False Then
        MsgBox "Faltan campos requeridos por informar.", vbExclamation, App.Title
        Exit Sub
    End If
    ' Hacemos el calculo si estan todos los requeridos
    Dim predijo As String
    Dim cadena As String
    Dim campo As String
    Dim Formula As String
    Dim Pos As Integer
    Dim ofor As New clsFormulas
    Dim encontrado As Boolean
    Dim j As Integer
'    prefijo = ""
'    Dim oDeter As New clsDeterminaciones
'    Dim oTD As New clsTipos_determinacion
'    oDeter.CargarDeterminacion (lista.ListItems(lista.SelectedItem.Index).Text)
'    oTD.CargarTipoDeterminacion (oDeter.getTIPO_DETERMINACION_ID)
    ofor.CARGAR (txtformula)
    cadena = ofor.getEXPRESION
    If Not IsNull(cadena) Then
        For i = 1 To Len(cadena)
            If Mid(cadena, i, 1) <> "C" Then
              If Mid(cadena, i, 1) = "," Then
                Formula = Formula & "."
              Else
                Formula = Formula & Mid(cadena, i, 1)
              End If
            Else
                Pos = InStr(i + 2, cadena, "_")
                campo = Mid(cadena, i + 2, (Pos) - (i + 2))
                j = datos.selectedItem.Index
                encontrado = False
                Do
                 If CInt(datos.ListItems(j).SubItems(3)) = CInt(campo) Then
                     Formula = Formula & Replace(datos.ListItems(j).SubItems(1), ",", ".")
                     encontrado = True
                 End If
                 j = j - 1
                Loop Until j = 0 Or encontrado = True
                i = Pos
            End If
        Next
    End If
    datos.ListItems(datos.selectedItem.Index).SubItems(1) = formatear(sc.Eval(Formula), datos.ListItems(datos.selectedItem.Index).SubItems(4), datos.ListItems(datos.selectedItem.Index).SubItems(5))
    grabar_auxdatos
    visualizar_duplicados
    pasar_siguiente_campo
    Exit Sub
fallo:
    MsgBox "Error en la formula. " & Err.Description, vbCritical, "Error"

End Sub


Private Sub cmdModificarEnsayo_Click()
   On Error GoTo cmdModificarEnsayo_Click_Error

    If cmbTipoEnsayo.getTEXTO = "" Then Exit Sub
    If MsgBox("�Esta seguro de modificar el tipo de ensato?", vbQuestion + vbYesNo, App.Title) = vbYes Then
        Dim oce_recepcion As New clsCe_recepcion
        oce_recepcion.InformarTipoEnsayo PK_ID_MUESTRA, cmbTipoEnsayo.getPK_SALIDA
        Set oce_recepcion = Nothing
        MsgBox "Ensayo modificado correctamente.", vbInformation, App.Title
    End If

   On Error GoTo 0
   Exit Sub

cmdModificarEnsayo_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure cmdModificarEnsayo_Click of Formulario frmCE_Resultados"
End Sub

Private Sub cmdModificarEquipo_Click()
   On Error GoTo cmdModificarEquipo_Click_Error

    If cmbEquipos.getPK_SALIDA <> 0 Then
        Dim i As Integer
        If txtusos = "" Then
            MsgBox "Debe indicar el n�mero de usos del equipo.", vbExclamation, App.Title
            Exit Sub
        End If
        If Not IsNumeric(txtusos) Then
            MsgBox "Debe indicar el n�mero de usos del equipo.", vbExclamation, App.Title
            Exit Sub
        End If
        Dim oEquipo As New clsEquipos
        oEquipo.Carga_Datos_Basicos cmbEquipos.getPK_SALIDA
        With listaEquipos.ListItems(listaEquipos.selectedItem.Index)
            .Text = oEquipo.getID_EQUIPO
            .SubItems(1) = oEquipo.getNOMBRE
            .SubItems(2) = oEquipo.getSERIE
            .SubItems(3) = 0
            .SubItems(4) = oEquipo.getNUMERO_USOS_MAXIMO
            .SubItems(5) = txtusos
        End With
        listaEquipos.ListItems(listaEquipos.selectedItem.Index).EnsureVisible
'        EQUIPOS_MODIFICADOS = True
        almacenar_equipos
        cmbEquipos.limpiar
        txtusos = ""
    End If

   On Error GoTo 0
   Exit Sub

cmdModificarEquipo_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure cmdModificarEquipo_Click of Formulario frmCE_Resultados"

End Sub

Private Sub cmdSpdaReiniciar_Click()
   On Error GoTo cmdSpdaReiniciar_Click_Error

    If MsgBox("�Esta seguro de reiniciar un nuevo SPDA?", vbQuestion + vbYesNo, App.Title) = vbYes Then
        Dim op As New clsParametros
        op.Carga parametros.SPDA_SECUENCIAL, ""
        op.setVALOR = CInt(op.getVALOR) + 1
        op.actualizar_valor parametros.SPDA_SECUENCIAL, ""
        op.setVALOR = "0"
        op.actualizar_valor parametros.SPDA_CANTIDAD, ""
                
        op.Carga parametros.SPDA_SECUENCIAL, ""
        frmSPDA.Caption = "SPDA N� : " & op.getVALOR
        op.Carga parametros.SPDA_CANTIDAD, ""
        txtSPDA.Text = op.getVALOR
    End If

   On Error GoTo 0
   Exit Sub

cmdSpdaReiniciar_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure cmdSpdaReiniciar_Click of Formulario frmCE_Resultados"
End Sub

Private Sub cmdVerificacion_Click()
   On Error GoTo cmdVerificacion_Click_Error

    If listaEquipos.ListItems.Count > 0 Then
        Dim objfrm  As New frmEquipoEdicionVerificacion
        Dim oEquipo As New clsEquipos
        oEquipo.Carga listaEquipos.ListItems(listaEquipos.selectedItem.Index).Text
        Set objfrm.EQUIPO = oEquipo
        
        If listaEquipos.ListItems(listaEquipos.selectedItem.Index).SubItems(3) = 0 Then
            
            objfrm.TipoEdicion = Alta
            objfrm.idVerificadorInternoInicial = USUARIO.getID_EMPLEADO
            objfrm.FechaProximaInicial = Now
        'MANTIS-810-I
        '   objfrm.IdPeriodoInicial = ENUM_EQUIPOS_PERIODICIDAD.ENUM_EQUIPOS_PERIODICIDAD_ANTES_ENSAYO
            objfrm.IdPeriodoInicial = ENUM_EQUIPOS_PERIODICIDAD.ENUM_EQUIPOS_PERIODICIDAD_ANTES_CADA_ENSAYO
        'MANTIS-810-F
            objfrm.IdTipoVerificacionIncial = 1
            
        'MANTIS-810-I
        '    objfrm.copiarUltimaVerificacionPeriodo = ENUM_EQUIPOS_PERIODICIDAD.ENUM_EQUIPOS_PERIODICIDAD_ANTES_ENSAYO
             objfrm.copiarUltimaVerificacionPeriodo = ENUM_EQUIPOS_PERIODICIDAD.ENUM_EQUIPOS_PERIODICIDAD_ANTES_ENSAYO
        'MANTIS-810-F
            objfrm.Show vbModal
          
            If objfrm.ID_VERIFICACION <> 0 Then
                listaEquipos.ListItems(listaEquipos.selectedItem.Index).SubItems(3) = objfrm.ID_VERIFICACION
            End If
            almacenar_equipos
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

   On Error GoTo 0
   Exit Sub

cmdVerificacion_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure cmdVerificacion_Click of Formulario frmCE_Resultados"
End Sub

Private Sub datos_Click()
   On Error GoTo datos_Click_Error

    On Error Resume Next
    If datos.ListItems.Count > 0 Then
        datos.selectedItem.EnsureVisible
        cmdCalcular.Enabled = False
        If datos.ListItems(datos.selectedItem.Index).bold = True Then
         If Trim(lblestado.Caption) = "" And datos.ListItems.Count > 1 Then
            cmdCalcular.Enabled = True
         Else
            If Trim(lblestado.Caption) = "DUPLICADA" And datos.ListItems.Count > 4 Then
                cmdCalcular.Enabled = True
                cmdCalcular_Click
                Exit Sub
            End If
         End If
        End If
        txtValor = Trim(datos.ListItems(datos.selectedItem.Index).SubItems(1))
        txtValor.SetFocus
        txtValor.SelStart = 0
        txtValor.SelLength = Len(txtValor)
        txtdato = datos.ListItems(datos.selectedItem.Index)
    End If

   On Error GoTo 0
   Exit Sub

datos_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure datos_Click of Formulario frmCE_Resultados"
End Sub

Private Sub Form_Initialize()
    log ("INITIALIZE : CE_RESULTADOS")
End Sub

Private Sub lista_Click()
   On Error GoTo lista_Click_Error

    If lista.ListItems.Count = 0 Then
        Exit Sub
    End If
    cargar_campos

   On Error GoTo 0
   Exit Sub

lista_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure lista_Click of Formulario frmCE_Resultados"
End Sub

Private Sub listaEquipos_Click()
   On Error GoTo listaEquipos_Click_Error

    If listaEquipos.ListItems.Count = 0 Then Exit Sub
    cmbEquipos.MostrarElemento listaEquipos.ListItems(listaEquipos.selectedItem.Index).Text
    txtusos = listaEquipos.ListItems(listaEquipos.selectedItem.Index).SubItems(5)

   On Error GoTo 0
   Exit Sub

listaEquipos_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure listaEquipos_Click of Formulario frmCE_Resultados"
End Sub

'Private Sub listaEquipos_ItemCheck(ByVal Item As MSComctlLib.ListItem)
'    EQUIPOS_MODIFICADOS = True
'End Sub

Private Sub txtvalor_GotFocus()
    txtValor.BackColor = &H80C0FF
    txtValor.SelStart = 0
    txtValor.SelLength = Len(Trim(txtValor))
End Sub

Private Sub txtvalor_KeyPress(KeyAscii As Integer)
    ' Escribir ',' al pulsar '.'
    If txtdato = "" Then
        Exit Sub
    End If
    If KeyAscii = 46 Then
         KeyAscii = 44
    End If
    On Error GoTo fallo
    
    If KeyAscii = 13 Then
        KeyAscii = 0
        If Trim(txtValor) = "" Or Trim(datos.ListItems(datos.selectedItem.Index).SubItems(3)) = "" Then
            datos.ListItems(datos.selectedItem.Index).SubItems(1) = " "
        Else
            datos.ListItems(datos.selectedItem.Index).SubItems(1) = formatear(txtValor, datos.ListItems(datos.selectedItem.Index).SubItems(5), datos.ListItems(datos.selectedItem.Index).SubItems(5))
        End If
        grabar_auxdatos
        visualizar_duplicados
        pasar_siguiente_campo
    End If
    
    Exit Sub
fallo:
    error_grave "Error en frmListadoDeterminaciones(txtvalor_KeyPress) : " & Err.Description

End Sub
Private Sub txtvalor_LostFocus()
    txtValor.BackColor = vbWhite
End Sub

Private Sub cmdAnadirEquipo_Click()
    If cmbEquipos.getPK_SALIDA <> 0 Then
        Dim i As Integer
        If txtusos = "" Then
            MsgBox "Debe indicar el n�mero de usos del equipo.", vbExclamation, App.Title
            Exit Sub
        End If
        If Not IsNumeric(txtusos) Then
            MsgBox "Debe indicar el n�mero de usos del equipo.", vbExclamation, App.Title
            Exit Sub
        End If
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
            .SubItems(3) = 0
            .SubItems(4) = oEquipo.getNUMERO_USOS_MAXIMO
            .SubItems(5) = txtusos
        End With
        listaEquipos.ListItems(listaEquipos.ListItems.Count).Checked = True
        listaEquipos.ListItems(listaEquipos.ListItems.Count).EnsureVisible
'        EQUIPOS_MODIFICADOS = True
        almacenar_equipos
        cmbEquipos.limpiar
        txtusos = ""
    End If
End Sub

Private Sub cmdEliminarEquipo_Click()
    If listaEquipos.ListItems.Count > 0 Then
'        EQUIPOS_MODIFICADOS = True
        listaEquipos.ListItems.Remove listaEquipos.selectedItem.Index
        almacenar_equipos
    End If
End Sub

Private Sub cmdObservador_Click()
Dim objfrm As New frmObservadorEnsayo

    objfrm.ES_CONTROL_EFICACIA = True
    objfrm.MUESTRA_ID = PK_ID_MUESTRA ' Id de la muestra
    objfrm.TIPO_DETERMINACION_ENSAYO_ID = mvarlngIdTipoEnsayo ' tipo del ensayo
    objfrm.DETERMINACION_ENSAYO_ID = 0
    objfrm.MUESTRA_CERRADA = mvarblnMuestra_Cerrada
    objfrm.TIPO_OBSERVACION_ID = MC_TIPOS_OBSERVACION.MCTO_CONTROL_EFICACIA
    
    objfrm.Show vbModal
    
    Set objfrm = Nothing
End Sub


Private Sub cmdPNT_Click()
    If IsNumeric(txttipoensayo) Then
        Dim oCE As New clsCe_tipos_ensayos
        oCE.Carga CLng(txttipoensayo)
        If oCE.getPNT_VINCULADO <> 0 Then
            Dim oPNT As New clsCa_documentos
            oPNT.mostrar oCE.getPNT_VINCULADO, True
            Set oPNT = Nothing
        Else
            MsgBox "El Tipo de Ensayo no tiene PNT V�nculado.", vbExclamation, App.Title
        End If
    End If
End Sub
Private Sub cmdAceptar_Click()
    Dim i As Integer
   On Error GoTo cmdaceptar_Click_Error
    ' Validar reactivos caducados (1090)
    Dim cont As Integer
    Dim existen As Boolean
    existen = False
    For cont = 1 To listaReactivos.ListItems.Count
        If Trim(listaReactivos.ListItems(cont).SubItems(2)) <> "" Then
            If Format(listaReactivos.ListItems(cont).SubItems(2), "yyyy-mm-dd") < Format(Date, "yyyy-mm-dd") Then
                existen = True
            End If
        End If
    Next
    If existen Then
        If MsgBox("Existen reactivos CADUCADOS. �ESTA SEGURO DE ALMACENAR LOS DATOS DE LA MUESTRA?", vbExclamation + vbYesNo, App.Title) = vbNo Then
            Exit Sub
        End If
    End If
    ' Verificar si es COC que introduzca la dimension
'    Dim oM As New clsMuestra
'    Dim oP As New clsParametros
'    oM.CargaMuestra PK_ID_MUESTRA
'    oP.Carga parametros.PARAM_HENKEL_TM, ""
'    If oM.getTIPO_MUESTRA_ID = oP.getVALOR Then
'        If cmbDimension.getTEXTO = "" Then
'            MsgBox "Debe indicar la dimensi�n de las probetas del ensayo.", vbCritical, App.Title
'            Exit Sub
'        End If
'    End If
    
    Dim oCe_resultados As New clsCe_resultados
    Dim RESULTADO As String
    If txtformula <> 0 Then
        almacenar_resultados_determinaciones
    Else
        For i = 0 To CInt(txtnumprobetas)
         If Not IsEmpty(xP(i, ColsP.Identificacion)) Then
            If CStr(xP(i, ColsP.Identificacion)) <> "" Then
                If chkEspesor.Value = Checked Then
                    If CStr(xP(i, ColsP.RESULTADO)) <> "" Then
                        Dim valores() As String
                        valores = Split(CStr(xP(i, ColsP.RESULTADO)), "-")
                        If UBound(valores) <> 2 Then
                            MsgBox "En los ensayos de espesor debe introducir los tres valores en el resultado. (Separados por - )", vbExclamation, App.Title
                            Exit Sub
                        End If
                        RESULTADO = valores(1)
                    End If
                Else
                    RESULTADO = CStr(xP(i, ColsP.RESULTADO))
                End If
                If RESULTADO <> "" Or Not IsEmpty(xP(i, ColsP.CONFORME)) Then
                    With oCe_resultados
                        If CStr(xP(i, ColsP.fecha)) <> "" Or CStr(xP(i, ColsP.RESULTADO)) <> "" Or xP(i, ColsP.CONFORME) = "Si" Or xP(i, ColsP.CONFORME) = "No" Then
                            If IsDate(CStr(xP(i, ColsP.fecha))) Then
                                .setFECHA = CStr(xP(i, ColsP.fecha))
                            Else
                                .setFECHA = Format(Date, "dd/mm/yyyy")
                            End If
                        Else
                            .setFECHA = ""
                        End If
                        .setRESULTADO = CStr(xP(i, ColsP.RESULTADO))
                        ' Conforme/No conforme
                        ' Si no esta marcado y el resultado si, hay que analizar los tangos
                        If xP(i, ColsP.CONFORME) = "Si" And RESULTADO = "" Then
                            .setCONFORME = 1
                        Else
                                If IsEmpty(xP(i, ColsP.CONFORME)) Then
                                    .setCONFORME = 1
                                Else
                                    If xP(i, ColsP.CONFORME) = "Si" Then
                                        .setCONFORME = 1
                                    Else
                                        .setCONFORME = 0
                                    End If
                                End If
                        End If
                        .Modificar_Resultado PK_ID_MUESTRA, CStr(xP(i, ColsP.DESIGNACION)), CStr(xP(i, ColsP.PROBETA)), CStr(xP(i, ColsP.AREA)), True
                    End With
                End If
            End If
         End If
        Next
    End If
    Dim oce_recepcion As New clsCe_recepcion
    With oce_recepcion
        If ddesde.Value = "01-01-1900" Then
            .setDURACION_FECHA_DESDE = ""
        Else
            .setDURACION_FECHA_DESDE = Format(ddesde.Value, "dd-mm-yyyy")
        End If
        If dhasta.Value = "01-01-1900" Then
            .setDURACION_FECHA_HASTA = ""
        Else
            .setDURACION_FECHA_HASTA = Format(dhasta.Value, "dd-mm-yyyy")
        End If
        If dhdesde.Value = "00:00:00" Then
            .setDURACION_HORA_DESDE = ""
        Else
            .setDURACION_HORA_DESDE = Format(dhdesde.Value, "hh:mm:ss")
        End If
        If dhhasta.Value = "00:00:00" Then
            .setDURACION_HORA_HASTA = ""
        Else
            .setDURACION_HORA_HASTA = Format(dhhasta.Value, "hh:mm:ss")
        End If
        If chkSinEspecificar.Value = Checked Then
'M1104            .setFECHA_PROCESADO_PIEZAS = ""
            .setFECHA_PROCESADO_PIEZAS = "NULL"
        Else
'M1104            .setFECHA_PROCESADO_PIEZAS = Format(fprocesado, "dd-mm-yyyy")
            .setFECHA_PROCESADO_PIEZAS = "'" & Format(fprocesado, "yyyy-mm-dd") & "'"
        End If
        .setESPESOR = txtEspesor
        ' Recorremos la lista de equipos
        Dim MAQUINA As String
        For i = 1 To listaEquipos.ListItems.Count
            MAQUINA = MAQUINA & listaEquipos.ListItems(i).Text & ";"
        Next
        .setMAQUINA = MAQUINA
        ' LOTE
        If cmbLote.getTEXTO <> "" Then
            .setLOTE_PROBETA_ID = cmbLote.getPK_SALIDA
        Else
            .setLOTE_PROBETA_ID = 0
        End If
        ' Reactivos
        Dim Reactivo As String
        Dim REACTIVOS_PROPIOS As String
        For i = 1 To listaReactivos.ListItems.Count
            If listaReactivos.ListItems(i).SubItems(3) = "E" Then
                Reactivo = Reactivo & listaReactivos.ListItems(i).Text & ";"
            End If
            If listaReactivos.ListItems(i).SubItems(3) = "I" Then
                REACTIVOS_PROPIOS = REACTIVOS_PROPIOS & listaReactivos.ListItems(i).Text & ";"
            End If
        Next
        .setREACTIVOS = Reactivo
        .setREACTIVOS_PROPIOS = REACTIVOS_PROPIOS
        ' Condiciones Ambientales
        .setCONDICIONES_AMBIENTALES = txtCondicionesAmbientales
        .setMATERIAL = txtMaterial
'        .setDIMENSION = txtDimension
        .setDIMENSION = cmbDimension.getPK_SALIDA
        .Informar_registro PK_ID_MUESTRA
    End With
    almacenar_equipos
    
    ' Verificar muestra cerrada
    Dim oMuestra As New clsMuestra
    oMuestra.comprobar_cierre (PK_ID_MUESTRA)
    MsgBox "Los datos se han guardado correctamente.", vbOKOnly + vbInformation, App.Title
    Unload Me
   On Error GoTo 0
   Exit Sub

cmdaceptar_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure cmdAceptar_Click of Formulario frmCE_Resultados2"
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
            .SubItems(1) = oTRPR.getCODIGO & "-" & Format(oRPR.getNUMERO, "000") & " " & oTRPR.getNOMBRE
            .SubItems(2) = Format(oRPR.getFECHA_CADUCIDAD, "dd-mm-yyyy")
            .SubItems(3) = "I"
        End With
        listaReactivos.ListItems(listaReactivos.ListItems.Count).EnsureVisible
    End If
    ' Limpiar Combos
    cmbReactivos.limpiar
    cmbReactivosInternos.limpiar
End Sub

Private Sub cmdEliminarReactivo_Click()
    If listaReactivos.ListItems.Count > 0 Then
        listaReactivos.ListItems.Remove listaReactivos.selectedItem.Index
        cmbReactivos.limpiar
        cmbReactivosInternos.limpiar
    End If
End Sub

Private Sub cmdImagen_Click()
    With frmCE_Imagenes
        .PK = PK_ID_MUESTRA
        .Show 1
    End With
End Sub

Private Sub cmdmodificarprobetas_Click()
    With frmCE_Recepcion_Probetas
        .PK_MUESTRA = PK_ID_MUESTRA
        .Show 1
        cargar_probetas
'        cargar_muestra
    End With
End Sub
Private Sub chkSinEspecificar_Click()
    If chkSinEspecificar.Value = Checked Then
        fprocesado.Value = "01/01/1900"
        fprocesado.Enabled = False
    Else
        fprocesado.Value = Date
        fprocesado.Enabled = True
    End If
End Sub

Private Sub cmdComienzo_Click()
    Dim s As String
   On Error GoTo cmdComienzo_Click_Error
    'M1281-I
    Dim fdesde_aux As Date
    Dim hdesde_aux As Date
    Dim fhasta_aux As Date
    Dim hhasta_aux As Date
    'M1281-F
'JGM    s = "�Establecer fechas del ensayo?  Se generar� un aviso de inicio y fin."
'JGM    If MsgBox(s, vbQuestion + vbYesNo, App.Title) = vbYes Then
        'M1281-I
        'ddesde = Date
        'dhdesde = Date & " " & Time
        fdesde_aux = Date
        hdesde_aux = Date & " " & Time
        'M1281-F
        Dim minuto As Integer
        If txtHoras <> "" Then
            minuto = InStr(1, txtHoras, ":")
            If minuto > 0 Then
                'M1281-I
                'dhhasta = DateAdd("h", Left(txtHoras, minuto - 1), dhdesde)
                'dhhasta = DateAdd("n", Mid(txtHoras, minuto + 1, Len(txtHoras) - minuto), dhhasta)
                hhasta_aux = DateAdd("h", Left(txtHoras, minuto - 1), hdesde_aux)
                hhasta_aux = DateAdd("n", Mid(txtHoras, minuto + 1, Len(txtHoras) - minuto), hhasta_aux)
                'M1281-F
            Else
                hhasta_aux = DateAdd("h", txtHoras, hdesde_aux)
            End If
            'M1281-I
            'dhasta = dhhasta
            'M1281-F
        End If
        'M1281-I
        If MsgBox("El Ensayo dar� comienzo el: " & vbCrLf & vbCrLf & Format(hdesde_aux, "DDDD dd/mm/yyyy a las hh:mm:ss") & vbCrLf & vbCrLf & " y finalizar� el: " & vbCrLf & vbCrLf & Format(hhasta_aux, "DDDD dd/mm/yyyy a las hh:mm:ss") & vbCrLf & vbCrLf & " �Desea continuar?", vbQuestion + vbYesNo, App.Title) = vbNo Then
            MsgBox "No se han establecido las fechas de ensayo", vbExclamation + vbOKOnly
            Exit Sub
        End If
        
        ddesde = fdesde_aux
        dhdesde = hdesde_aux
        dhasta = hhasta_aux
        dhhasta = hhasta_aux
        'M1281-F
        ' Enviar aviso
        Dim oMensaje As New clsMensajes
        Dim oMuestra As New clsMuestra
        oMuestra.CargaMuestra (PK_ID_MUESTRA)
        Dim mens As Integer
        With oMensaje
            .setASUNTO = Trim(str(oMuestra.getID_GENERAL)) & " (" & oMuestra.CodigoParticular(gmuestra) & ")" & " Finalizaci�n de Control de eficacia"
            .setTEXTO = .getTEXTO & "El usuario " & USUARIO.getUSUARIO & " ha iniciado un control de eficacia. " & vbNewLine & vbNewLine
            .setTEXTO = .getTEXTO & "Fecha de comienzo : " & dhdesde & vbNewLine & vbNewLine
            .setTEXTO = .getTEXTO & "Fecha de finalizaci�n : " & dhhasta & vbNewLine
            .setEMPLEADO_ID = USUARIO.getID_EMPLEADO
'            .setFECHA_INICIO = Format(ddesde.value, "yyyy-mm-dd")
            .setFECHA_INICIO = Format(dhhasta.Value, "yyyy-mm-dd")
            .setFECHA_FIN = Format(dhhasta.Value, "yyyy-mm-dd")
            
            .setACCION = "frmVerMuestra;" & PK_ID_MUESTRA
            .setHORA_INICIO = Format(dhhasta.Value, "hh:mm:ss")
            .setHORA_FIN = Format(dhhasta.Value, "hh:mm:ss")
            .setCATEGORIA = MENSAJES_CATEGORIAS.MENSAJES_CATEGORIAS_CE
            .setDURACION = 0
            
            mens = .Insertar
            If mens > 0 Then
                Dim omu As New clsMensajes_usuarios
                Dim i As Integer
                Dim usuarios() As String
                Dim opar As New clsParametros
                If (opar.Carga(11, "")) Then
                    usuarios = Split(opar.getVALOR, ",")
                    For i = LBound(usuarios) To UBound(usuarios)
                        If usuarios(i) <> "" Then
                            omu.setEMPLEADO_ID = usuarios(i)
                            omu.setMENSAJE_ID = mens
                            omu.Insertar
                        End If
                    Next
                End If
                frmCalendario.cargar_eventos
            End If
        End With
        Dim oce_recepcion As New clsCe_recepcion
        With oce_recepcion
            .setDURACION_FECHA_DESDE = Format(ddesde.Value, "dd-mm-yyyy")
            .setDURACION_HORA_DESDE = Format(dhdesde.Value, "hh:mm:ss")
            .setDURACION_FECHA_HASTA = Format(dhhasta.Value, "dd-mm-yyyy")
            .setDURACION_HORA_HASTA = Format(dhhasta.Value, "hh:mm:ss")
            .Informar_Duracion_Ensayo PK_ID_MUESTRA
        End With
        MsgBox "Fechas establecidas correctamente.", vbInformation, App.Title
'JGM    End If

   On Error GoTo 0
   Exit Sub

cmdComienzo_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure cmdComienzo_Click of Formulario frmCE_Resultados2"
End Sub

Private Sub cmdSalir_Click()
'    If (lblCerrada = "ABIERTA" And EQUIPOS_MODIFICADOS = True) Or chkModificar.value = Checked Then
'        almacenar_equipos
'    End If
    Unload Me
End Sub

Private Sub cmdtipoensayo_Click()
    If IsNumeric(txttipoensayo) Then
        frmCE_Tipo_Ensayo.PK = CLng(txttipoensayo)
        frmCE_Tipo_Ensayo.Show 1
        cargar_datos_tipo_ensayo CLng(txttipoensayo)
    End If
End Sub

'Private Sub Form_Activate()
'    If blnPrimeraVez Then
'        gridP_BeforeColEdit ColsP.RESULTADO, 0, 0
'        blnPrimeraVez = False
'    End If
'End Sub

Private Sub Form_Load()
   On Error GoTo Form_Load_Error

    log (Me.Name)
    log "CE_RESULTADOS - cargar_botones"
    cargar_botones Me
    log "CE_RESULTADOS - permisos"
    permisos
    log "CE_RESULTADOS - cabecera"
    cabecera
    log "CE_RESULTADOS - cargar_combos"
    cargar_combos
    
    If PK_ID_MUESTRA > 0 Then
        log "CE_RESULTADOS - cargar_muestra"
        cargar_muestra
    End If
    
'    blnPrimeraVez = False
'    EQUIPOS_MODIFICADOS = False
    log "CE_RESULTADOS - Fin Load"
    
'    Call ConfigurarTablet

   On Error GoTo 0
   Exit Sub

Form_Load_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure Form_Load of Formulario frmCE_Resultados"
    
End Sub
Private Sub cargar_combos()
   On Error GoTo cargar_combos_Error

    llenar_combo cmbTipoEnsayo, New clsCe_tipos_ensayos, 0, frmCE_Tipo_Ensayo, ""
    llenar_combo cmbLote, New clsCe_lotes_probetas, 0, frmCE_Lote_Probeta, ""
    cmbLote.desactivar
    llenar_combo cmbEquipos, New clsEquipos, 0, frmEquipoEdicion, ""
    llenar_combo cmbReactivos, New clsBotes_ex, 0, Me, "AND ABIERTO = 1 AND FINALIZADO = 0 "
    llenar_combo cmbReactivosInternos, New clsRpr_botes, 0, Me, " AND ISNULL(fecha_fin)"
    Dim oDeco As New clsDecodificadora
    oDeco.cargar_mi_combo cmbDimension, DECODIFICADORA.DECODIFICADORA_DIMENSIONES

   On Error GoTo 0
   Exit Sub

cargar_combos_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure cargar_combos of Formulario frmCE_Resultados"
End Sub

Private Sub cargar_muestra()
    'Titulo
    Dim oMuestra As New clsMuestra
   On Error GoTo cargar_muestra_Error

    log "CE_RESULTADOS - cargar_muestra - CargaMuestra"
    oMuestra.CargaMuestra (PK_ID_MUESTRA)
    lbltitulo = "Registro resultados muestra : " & Trim(str(oMuestra.getID_GENERAL)) & " (" & oMuestra.CodigoParticular(gmuestra) & ")"
    Me.Caption = lbltitulo
    ' Duplicada
    If oMuestra.getANALISIS_DUPLICADO = 1 Then
        chkDuplicada.Value = Checked
    End If
    ' SPDA
    frmSPDA.visible = False
    Dim oParametros As New clsParametros
    oParametros.Carga parametros.SPDA_TIPO_MUESTRA, ""
    If oParametros.getVALOR <> "" Then
        If oParametros.getVALOR = oMuestra.getTIPO_MUESTRA_ID Then
            frmSPDA.visible = True
            oParametros.Carga parametros.SPDA_SECUENCIAL, ""
            frmSPDA.Caption = "SPDA N� : " & oParametros.getVALOR
            oParametros.Carga parametros.SPDA_CANTIDAD, ""
            txtSPDA.Text = oParametros.getVALOR
        End If
    End If
    log "CE_RESULTADOS - cargar_muestra - ESP"
    ' Verificar si es un ensayo de espesor
    If oMuestra.getTIPO_MUESTRA_ID = TIPOS_MUESTRAS.ESPESOR Or _
       oMuestra.getTIPO_MUESTRA_ID = TIPOS_MUESTRAS.MICRODUREZA Or _
       oMuestra.getTIPO_MUESTRA_ID = TIPOS_MUESTRAS.RUGOSIDAD Then
        chkEspesor.Value = Checked
        lblMensaje = "ESP : Introduzca N�Medidas,Media y Desviaci�n separados por gui�n (-)"
    Else
        lblMensaje = ""
        chkEspesor.Value = Unchecked
    End If
    ' Si es un ECS, habilitar los botones para introducir las maquinas
    If oMuestra.getTIPO_MUESTRA_ID = TIPOS_MUESTRAS.ecs Then
        frmEquipos.Caption = "Equipos Carga Sostenida (M�quina-C�lula-Indicador)"
    End If
    ' CE
    Dim oce_recepcion As New clsCe_recepcion
    log "CE_RESULTADOS - cargar_muestra - oce_recepcion"
    If oce_recepcion.Carga(PK_ID_MUESTRA) Then
        txttipoensayo = oce_recepcion.getTIPO_ENSAYO_ID
        cargar_datos_tipo_ensayo oce_recepcion.getTIPO_ENSAYO_ID
        mvarlngIdTipoEnsayo = oce_recepcion.getTIPO_ENSAYO_ID
        Dim oTipo_ensayo As New clsCe_tipos_ensayos
        oTipo_ensayo.Carga (oce_recepcion.getTIPO_ENSAYO_ID)
        ' CE001
        ' Condiciones Ambientales
        txtCondicionesAmbientales = oce_recepcion.getCONDICIONES_AMBIENTALES
        txtMaterial = oce_recepcion.getMATERIAL
'        txtDimension = oce_recepcion.getDIMENSION
        cmbDimension.MostrarElemento oce_recepcion.getDIMENSION
        ' CE por formula o Resultado de probetas
        If oTipo_ensayo.getFORMULA_ID = 0 Then
            gridP.visible = True
            frmResultados.visible = False
            
            txtformula = "0"
        Else
            gridP.visible = False
            frmResultados.visible = True
            ' EN LUGAR DE CARGAR LOS CAMPOS DE LA FORMULA DEL ENSAYO, MIRAMOS SI EXISTE YA ALGUN RESULTADO INSERTADO, EN CUYO CASO LA RECUPERAMOS EN FUNCION AL CAMPO
            Dim oCERD As New clsCe_resultados_determinaciones
            txtformula = oCERD.recuperarFormula(PK_ID_MUESTRA)
            If txtformula = "" Then
                txtformula = oTipo_ensayo.getFORMULA_ID
            End If
        End If
        ' Resto de datos
        If oTipo_ensayo.getINCLUYE_ESPESOR = 1 Then
            txtEspesor = oce_recepcion.getESPESOR
        End If
        If CInt(oTipo_ensayo.getLOTE_PROBETAS) = 1 Then
            cmbLote.MostrarElemento oce_recepcion.getLOTE_PROBETA_ID
        End If
        
        Dim oEquipo As New clsEquipos
        log "CE_RESULTADOS - cargar_muestra - oce_recepcion - Inicio"
        With oce_recepcion
            Dim i As Integer
'            If .getMAQUINA <> "" Then
                cargar_equipos PK_ID_MUESTRA
'            End If
            If .getFECHA_PROCESADO_PIEZAS = "" Then
                chkSinEspecificar.Value = Checked
            Else
                fprocesado = Format(.getFECHA_PROCESADO_PIEZAS, "dd-mm-yyyy")
            End If
            If .getDURACION_FECHA_DESDE = "" Then
                ddesde = "01-01-1900"
            Else
                ddesde = Format(.getDURACION_FECHA_DESDE, "dd-mm-yyyy")
            End If
            If .getDURACION_FECHA_HASTA = "" Then
                dhasta = "01-01-1900"
            Else
                dhasta = Format(.getDURACION_FECHA_HASTA, "dd-mm-yyyy")
            End If
            If .getDURACION_HORA_DESDE <> "" Then
                dhdesde.Value = Date & " " & .getDURACION_HORA_DESDE
            End If
            If .getDURACION_HORA_HASTA <> "" Then
                dhhasta.Value = Date & " " & .getDURACION_HORA_HASTA
            End If
            ' REACTIVOS EXTERNOS
            log "CE_RESULTADOS - cargar_muestra - oce_recepcion - REX"
            listaReactivos.ListItems.Clear
            If .getREACTIVOS <> "" Then
                Dim REACTIVOS() As String
                Dim oReactivo As New clsBotes_ex
                Dim oTb As New clsTipos_bote_ex
                Dim oTR As New clsTipos_reactivo_ex
                REACTIVOS = Split(.getREACTIVOS, ";")
                For i = LBound(REACTIVOS) To UBound(REACTIVOS) - 1
                    oReactivo.CARGAR CLng(REACTIVOS(i))
                    oTb.CARGAR oReactivo.getTIPO_BOTE_EX_ID
                    oTR.CARGAR oTb.getTIPO_REACTIVO_EX_ID
                    With listaReactivos.ListItems.Add(, , REACTIVOS(i))
                        .SubItems(1) = oTR.getNOMBRE
                        .SubItems(2) = Format(oReactivo.getFECHA_CADUCIDAD, "DD-MM-YYYY")
                        .SubItems(3) = "E"
                    End With
                Next
            End If
            ' REACTIVOS PROPIOS
            log "CE_RESULTADOS - cargar_muestra - oce_recepcion - REX PROPIOS"
            If .getREACTIVOS_PROPIOS <> "" Then
                Dim REACTIVOS_PROPIOS() As String
                Dim oRPR As New clsRpr_botes
                Dim oTRPR As New clsRPR_Tipos
                REACTIVOS_PROPIOS = Split(.getREACTIVOS_PROPIOS, ";")
                For i = LBound(REACTIVOS_PROPIOS) To UBound(REACTIVOS_PROPIOS) - 1
                    oRPR.Carga CLng(REACTIVOS_PROPIOS(i))
                    oTRPR.CARGAR oRPR.getTIPO_REACTIVO_PR_ID
                    With listaReactivos.ListItems.Add(, , REACTIVOS_PROPIOS(i))
                        .SubItems(1) = oTRPR.getCODIGO & "-" & Format(oRPR.getNUMERO, "000") & " " & oTRPR.getNOMBRE
                        .SubItems(2) = Format(oRPR.getFECHA_CADUCIDAD, "DD-MM-YYYY")
                        .SubItems(3) = "I"
                    End With
                Next
            End If
        End With
        log "CE_RESULTADOS - cargar_muestra - oce_recepcion - cargar_probetas"
        cargar_probetas
    End If
    log "CE_RESULTADOS - cargar_muestra - oce_recepcion - proteger_campos"
    proteger_campos oMuestra.getCERRADA
    
    Set oce_recepcion = Nothing

   On Error GoTo 0
   Exit Sub

cargar_muestra_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure cargar_muestra of Formulario frmCE_Resultados2"
End Sub

Private Sub cargar_probetas()
    Dim oCe_resultados As New clsCe_resultados
    Dim rs As ADODB.Recordset
   On Error GoTo cargar_probetas_Error

    Set rs = oCe_resultados.Listado_por_muestra(PK_ID_MUESTRA)
    Dim i As Integer
    If txtformula = 0 Then
        i = 0
        If rs.RecordCount = 1 Then ' Cuando solo tiene una probeta, hay que meter 2 filas para que pueda actualizar
            inicializar_grid rs.RecordCount
            txtnumprobetas = rs.RecordCount
        Else
            inicializar_grid rs.RecordCount - 1
            txtnumprobetas = rs.RecordCount - 1
        End If
        If rs.RecordCount > 0 Then
            Do
                xP(i, ColsP.Identificacion) = CStr(rs("IDENTIFICACION_CANAGROSA"))
                xP(i, ColsP.IDENTIFICACION_CLIENTE) = CStr(rs("IDENTIFICACION_CLIENTE"))
                xP(i, ColsP.DIMENSION) = CStr(rs("DIMENSION"))
                If rs("FECHA") <> "" Then
                    xP(i, ColsP.fecha) = CStr(rs("FECHA"))
                    xP(i, ColsP.RESULTADO) = CStr(rs("RESULTADO"))
                    If rs("CONFORME") = 0 Then
                        xP(i, ColsP.CONFORME) = CStr("No")
                    Else
                        xP(i, ColsP.CONFORME) = CStr("Si")
                    End If
                End If
                xP(i, ColsP.DESIGNACION) = CStr(rs("DESIGNACION"))
                xP(i, ColsP.PROBETA) = CStr(rs("PROBETA"))
                xP(i, ColsP.AREA) = CStr(rs("AREA"))
                i = i + 1
                rs.MoveNext
            Loop Until rs.EOF
        End If
        Set oCe_resultados = Nothing
        gridP.Refresh
        gridP.Rebind
        gridP.Col = ColsP.RESULTADO
    Else
        lista.ListItems.Clear
        txtnumprobetas = rs.RecordCount
        If rs.RecordCount > 0 Then
            Do
                With lista.ListItems.Add(, , rs("DESIGNACION"))
                    .SubItems(1) = rs("PROBETA")
                    .SubItems(2) = rs("AREA")
                    .SubItems(3) = rs("IDENTIFICACION_CANAGROSA")
                    .SubItems(4) = rs("IDENTIFICACION_CLIENTE")
                    .SubItems(5) = Format(rs("FECHA"), "dd-mm-yyyy") & ""
                    If rs("RESULTADO") <> "" Then
                        .SubItems(6) = rs("RESULTADO") & ""
                    Else
                        .SubItems(6) = " "
                    End If
                End With
                rs.MoveNext
            Loop Until rs.EOF
            lista_Click
        End If
    End If
    Set rs = Nothing

   On Error GoTo 0
   Exit Sub

cargar_probetas_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure cargar_probetas of Formulario frmCE_Resultados"
End Sub
Private Sub inicializar_grid(filas)
   On Error GoTo inicializar_grid_Error

    gridP.Col = 0
    gridP.Row = 0
    xP.Clear
    xP.ReDim 0, filas, 0, ColP
'    xP.Clear
    Set gridP.Array = xP
    gridP.Refresh

   On Error GoTo 0
   Exit Sub

inicializar_grid_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure inicializar_grid of Formulario frmCE_Recepcion_Nuevo2"
End Sub
Private Sub fprocesado_Change()
   On Error GoTo fprocesado_Change_Error

    chkSinEspecificar.Value = Unchecked

   On Error GoTo 0
   Exit Sub

fprocesado_Change_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure fprocesado_Change of Formulario frmCE_Resultados"
End Sub
Private Sub gridP_AfterColEdit(ByVal ColIndex As Integer)
'    On Error Resume Next
    Dim fila As String
   On Error GoTo gridP_AfterColEdit_Error

    fila = gridP.Row
    Dim CELDA As String
    CELDA = gridP.Text
'    MsgBox FILA & ":" & xP(FILA, ColsP.Resultado) & ":" & gridP.Text
    Select Case ColIndex
        Case ColsP.RESULTADO
            If CELDA <> "" Then
              ' Recuperamos el valor de la media para el espesor o tomamos el resultado
              Dim VALOR As String
              If chkEspesor.Value = Checked Then
                Dim valores() As String
                valores = Split(CELDA, "-")
                Dim i As Integer
                If UBound(valores) <> 2 Then
                    MsgBox "Para los ensayos de espesor, debe introducir los tres valores.", vbExclamation, App.Title
                    
                    Exit Sub
                End If
                VALOR = valores(1)
              Else
                VALOR = CELDA
              End If
              ' Validar el resultados con los rangos
              If IsNumeric(VALOR) Then
                If IsNumeric(txtDatos(3)) Then
                  If CSng(txtDatos(3)) > CSng(VALOR) Then
                     MsgBox "ATENCION: El valor introducido es MENOR que el m�nimo establecido.", vbExclamation, App.Title
                  End If
                End If
                If IsNumeric(txtDatos(4)) Then
                  If CSng(txtDatos(4)) < CSng(VALOR) Then
                     MsgBox "ATENCION: El valor introducido es MAYOR que el m�ximo establecido.", vbExclamation, App.Title
                  End If
                End If
              End If
            End If
            If gridP.Row = CInt(txtnumprobetas) Then
                gridP.Row = 0
            End If
        Case ColsP.RESULTADO, ColsP.CONFORME, ColsP.fecha
            If gridP.Row = CInt(txtnumprobetas) Then
                gridP.Row = 0
            Else
                gridP.Row = gridP.Row + 1
            End If
    End Select

   On Error GoTo 0
   Exit Sub

gridP_AfterColEdit_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure gridP_AfterColEdit of Formulario frmCE_Resultados"
End Sub


'Private Sub gridP_BeforeColEdit(ByVal ColIndex As Integer, ByVal KeyAscii As Integer, Cancel As Integer)
'
'   On Error GoTo gridP_BeforeColEdit_Error
'
'    If blnEsTablet And ColIndex = ColsP.RESULTADO Then
'        gridP.Col = ColIndex
'        If Trim(gridP.Text) <> "" Then
'            TecladoNumerico.TextoInicial = gridP.Text
'            TecladoNumerico.cabecera = xP(gridP.Row, 0)
'            TecladoNumerico.Subcabecera = xP(gridP.Row, 1)
'            If Trim(xP(gridP.Row, ColsP.CONFORME)) = "" Then
'                TecladoNumerico.CONFORME = -1
'            ElseIf Trim(xP(gridP.Row, ColsP.CONFORME)) <> "Si" Then
'                TecladoNumerico.CONFORME = 1
'            Else
'                TecladoNumerico.CONFORME = 0
'            End If
'            TecladoNumerico.Show 1
'            gridP.EditActive = False
'        End If
'    End If
'
'   On Error GoTo 0
'   Exit Sub
'
'gridP_BeforeColEdit_Error:
'
'    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure gridP_BeforeColEdit of Formulario frmCE_Resultados"
'
'End Sub


Private Sub gridP_KeyPress(KeyAscii As Integer)
    If gridP.Col = ColsP.RESULTADO Then
        If KeyAscii = 46 Then
             KeyAscii = 44
        End If
    End If
End Sub

Private Sub cabecera()
   On Error GoTo cabecera_Error

    With listaEquipos.ColumnHeaders
        .Add , , "N�Equipo", 800, lvwColumnLeft
        .Add , , "Nombre", 3000, lvwColumnLeft
        .Add , , "N�Serie", 1000, lvwColumnCenter
        .Add , , "Verificacion", 1, lvwColumnCenter
        .Add , , "Usos M�x", 700, lvwColumnCenter
        .Add , , "N� Usos", 700, lvwColumnCenter
    End With
    With listaReactivos.ColumnHeaders
        .Add , , "N�mero", 800, lvwColumnLeft
        .Add , , "Reactivo", 3200, lvwColumnLeft
        .Add , , "Caducidad", 1200, lvwColumnCenter
        .Add , , "TIPO", 0, lvwColumnCenter ' (I-E) Interno o externo
    End With
    ' Listas para formulas
    With lista.ColumnHeaders
        .Add , , "DESIGNACION", 1, lvwColumnLeft
        .Add , , "PROBETA", 1, lvwColumnLeft
        .Add , , "AREA", 1, lvwColumnLeft
        .Add , , "Identificaci�n Canagrosa", 2350, lvwColumnLeft
        .Add , , "Identificaci�n Cliente", 2350, lvwColumnLeft
        .Add , , "Fecha", 1100, lvwColumnCenter
        .Add , , "Resultado", 1000, lvwColumnRight
    End With
    ' Datos
    With datos.ColumnHeaders
        .Add , , "Campo", 3000, lvwColumnLeft
        .Add , , "Valor", 1500, lvwColumnRight
        .Add , , "Unidad", 1000, lvwColumnLeft
        .Add , , "ID", 700, lvwColumnCenter
        .Add , , "Enteros", 0, lvwColumnCenter
        .Add , , "Decimales", 0, lvwColumnCenter
    End With
    ' Aux Datos
    With auxdatos.ColumnHeaders
        .Add , , "DESIGNACION", 1, lvwColumnLeft
        .Add , , "PROBETA", 1, lvwColumnLeft
        .Add , , "AREA", 1, lvwColumnLeft
        .Add , , "Valor", 1000, lvwColumnLeft
        .Add , , "Linea", 1000, lvwColumnLeft
        .Add , , "Campo", 1000, lvwColumnLeft
        .Add , , "Media", 200, lvwColumnLeft
    End With

   On Error GoTo 0
   Exit Sub

cabecera_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure cabecera of Formulario frmCE_Resultados"
End Sub


Private Sub listaEquipos_DblClick()
    If listaEquipos.ListItems.Count > 0 Then
        frmEquipoEdicion.PK = listaEquipos.ListItems(listaEquipos.selectedItem.Index)
        frmEquipoEdicion.Show 1
    End If
End Sub

'Private Sub TecladoNumerico_Change(ByVal res As String)
'    gridP.Text = res
'End Sub
'
'
'Private Sub ConfigurarTablet()
'
'    blnEsTablet = pc_es_tablet
'    If blnEsTablet Then
'        Set TecladoNumerico = New frmTecladoNumerico
'        TecladoNumerico.posX = Screen.Width - TecladoNumerico.Width
'        TecladoNumerico.posY = 0
'        blnPrimeraVez = True
'        gridP.Columns(ColsP.RESULTADO).Locked = True
'        Me.top = 0
'    End If
'End Sub
'
'Private Sub TecladoNumerico_EstablecerConformidad(ByVal VALOR As Integer)
'    If VALOR > -1 Then
'
'        gridP.Columns(ColsP.CONFORME) = IIf(VALOR = 1, "Si", "No")
'
'    Else
'        gridP.Columns(ColsP.CONFORME) = ""
'    End If
'
'    gridP.Col = ColsP.RESULTADO
'End Sub
'
'
'Private Sub TecladoNumerico_SiguienteElemento(cabecera As String, Subcabecera As String, RESULTADO As String, fecha As String, CONFORME As Integer, Cerrar As Boolean, desestimarEvento As Boolean)
'If gridP.Row + 1 > filasP Then
'    TecladoNumerico.Hide
'    gridP.EditActive = False
'    Exit Sub
'End If
'
'' si existe siguiente Fila, edita la siguiente fila
'
'If (gridP.Row + 1) <= xP.UpperBound(1) Then
'    If Not IsEmpty(xP(gridP.Row + 1, 0)) Then
'        If Trim(xP(gridP.Row + 1, 0)) <> "" Then
'            gridP.EditActive = False
'            gridP.Row = gridP.Row + 1
'            RESULTADO = gridP.Text
'            cabecera = xP(gridP.Row, 0)
'            Subcabecera = xP(gridP.Row, 1)
'            fecha = xP(gridP.Row, 1)
'            gridP.EditActive = True
'        End If
'    ElseIf txtnumprobetas.Text = "1" Then
'        gridP.Row = 1
'        Cerrar = True
'        gridP.EditActive = False
'    End If
'Else
'    If txtnumprobetas.Text = "0" Then
'        gridP.Row = 1
'    Else
'        gridP.Row = 0
'    End If
'
'    Cerrar = True
'    gridP.EditActive = False
'End If
'End Sub

Private Sub cargar_datos_tipo_ensayo(ENSAYO As Long)

    Dim oTipo_ensayo As New clsCe_tipos_ensayos
   On Error GoTo cargar_datos_tipo_ensayo_Error

    If oTipo_ensayo.Carga(ENSAYO) Then
'            If IsNumeric(oTipo_ensayo.getHORAS) Then
            If oTipo_ensayo.getHORAS <> "" Then
                Frame4.visible = True
                txtHoras = oTipo_ensayo.getHORAS
            Else
                Frame4.visible = False
                txtHoras = ""
                txtHoras.visible = False
            End If
            If oTipo_ensayo.getINCLUYE_ESPESOR = 1 Then
'                txtespesor = oce_recepcion.getESPESOR
                txtEspesor.Enabled = True
                txtEspesor.BackColor = vbWhite
            Else
                txtEspesor = "No requiere espesor."
                txtEspesor.Enabled = False
            End If
            If CInt(oTipo_ensayo.getLOTE_PROBETAS) = 1 Then
                chkLote.Value = Checked
                cmbLote.activar
 '               cmbLote.MostrarElemento oce_recepcion.getLOTE_PROBETA_ID
            Else
                chkLote.Value = Unchecked
            End If
    
'            txtDatos(0) = oTipo_ensayo.getNOMBRE
            cmbTipoEnsayo.MostrarElemento oTipo_ensayo.getID_TIPO_ENSAYO
            txtDatos(1) = oTipo_ensayo.getCRITERIO
            txtDatos(2) = oTipo_ensayo.getNORMA
            txtDatos(3) = oTipo_ensayo.getRANGO_MIN
            txtDatos(4) = oTipo_ensayo.getRANGO_MAX
            If oTipo_ensayo.getPNT_VINCULADO <> 0 Then
                Dim oDoc As New clsCa_documentos
                oDoc.Carga oTipo_ensayo.getPNT_VINCULADO
                txtDatos(5) = oDoc.getNOMBRE
            End If
    End If

   On Error GoTo 0
   Exit Sub

cargar_datos_tipo_ensayo_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure cargar_datos_tipo_ensayo of Formulario frmCE_Resultados"
End Sub
Private Sub proteger_campos(CERRADA As Integer)
   On Error GoTo proteger_campos_Error

    If CERRADA = 1 Then
        cmdComienzo.Enabled = False
        cmbTipoEnsayo.desactivar
        cmdModificarEnsayo.Enabled = False
        
        If chkModificar.Value = Unchecked Then
            cmdAceptar.Enabled = False
            
            cmdEliminarReactivo.Enabled = False
            cmdAnadirReactivo.Enabled = False
            cmdEliminarEquipo.Enabled = False
            cmdAnadirEquipo.Enabled = False
            cmbEquipos.desactivar
            cmbReactivos.desactivar
            cmbReactivosInternos.desactivar
        
            frmFechasEnsayo.Enabled = False
        
        Else
            cmdAceptar.Enabled = True
            
            cmdEliminarReactivo.Enabled = True
            cmdAnadirReactivo.Enabled = True
            cmdEliminarEquipo.Enabled = True
            cmdAnadirEquipo.Enabled = True
            cmbEquipos.activar
            cmbReactivos.activar
            cmbReactivosInternos.activar
            
            frmFechasEnsayo.Enabled = True
        End If
' FIN-LORENA

        gridP.EditActive = False
        
        mvarblnMuestra_Cerrada = True
        
    Else
        cmdAceptar.Enabled = True
        cmdComienzo.Enabled = True
        cmbTipoEnsayo.activar
        cmdModificarEnsayo.Enabled = True

        cmdEliminarReactivo.Enabled = True
        cmdAnadirReactivo.Enabled = True
        cmdEliminarEquipo.Enabled = True
        cmdAnadirEquipo.Enabled = True
        cmbEquipos.activar
        cmbReactivos.activar
        cmbReactivosInternos.activar
        frmFechasEnsayo.Enabled = True
        gridP.EditActive = True
        
        mvarblnMuestra_Cerrada = False
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

   On Error GoTo 0
   Exit Sub

proteger_campos_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure proteger_campos of Formulario frmCE_Resultados"
End Sub
Private Sub siguiente_campo()
   On Error GoTo siguiente_campo_Error

    If lista.ListItems.Count > lista.selectedItem.Index Then
        Set lista.selectedItem = lista.ListItems(lista.selectedItem.Index + 1)
        lista_Click
        datos_Click
    Else
        datos.ListItems.Clear
        txtdato = ""
        txtValor = ""
        datos.SetFocus
    End If

   On Error GoTo 0
   Exit Sub

siguiente_campo_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure siguiente_campo of Formulario frmCE_Resultados"
End Sub
Private Sub cargar_campos()
    Dim ocampos As New clsFormulas_campos
    Dim rs As New ADODB.Recordset
    Dim consulta As String
    Dim duplicado As Integer
    Dim nombre As String
    Dim i As Integer
    Dim j As Integer
   On Error GoTo cargar_campos_Error

    datos.ListItems.Clear
    cmdCalcular.Enabled = False
    Set rs = ocampos.ListaFormulas(txtformula)
    Label5(0).Width = 6465
    lblestado.Caption = ""
    If chkDuplicada.Value = Checked Then
        duplicado = 2
        Label5(0).Width = 3900
        lblestado.Caption = "DUPLICADA"
    Else
        duplicado = 1
    End If
    Dim rs_campos As ADODB.Recordset
    Dim oCE_RD As New clsCe_resultados_determinaciones
    If rs.RecordCount <> 0 Then
     For j = 1 To duplicado
      rs.MoveFirst
      While Not rs.EOF
        Set rs_campos = ocampos.CampoConUnidad(rs("id_campo"))
        If rs_campos.RecordCount > 0 Then
            If duplicado = 2 Then
                nombre = rs_campos(0) & " (" & j & ")"
            Else
                nombre = rs_campos(0)
            End If
            With datos.ListItems.Add(, , nombre)
                    .SubItems(1) = " "
                    If oCE_RD.Carga(PK_ID_MUESTRA, lista.ListItems(lista.selectedItem.Index).Text, lista.ListItems(lista.selectedItem.Index).SubItems(1), lista.ListItems(lista.selectedItem.Index).SubItems(2), rs("id_campo")) Then
                      If j = 1 Then
                        .SubItems(1) = Replace(oCE_RD.getVALOR_1, ".", ",")
                      Else
                        .SubItems(1) = Replace(oCE_RD.getVALOR_2, ".", ",")
                      End If
                    End If
                    .SubItems(2) = rs_campos(1)
                    .SubItems(3) = rs_campos(2)
                    .SubItems(4) = rs_campos(4) ' ENTEROS
                    .SubItems(5) = rs_campos(5) ' DECIMALES
                End With
            If rs_campos(3) <> 0 Then ' ES_SOLUCION
                datos.ListItems.Item(datos.ListItems.Count).bold = True
            End If
        End If
        rs.MoveNext
      Wend
     Next
     ' Resultados duplicados
     If duplicado = 2 Then
        With datos.ListItems.Add(, , "Resultado (MEDIA)")
            .SubItems(1) = " "
        End With
        With datos.ListItems.Add(, , "Dif. entre duplicados")
            .SubItems(1) = " "
        End With
        'M1371-I
        'With datos.ListItems.Add(, , "Revisi�n de duplicados")
        '    .SubItems(1) = " "
        'End With
        'M1371-F
     End If
     visualizar_duplicados
    End If
    ' Comprobar si ya tiene datos
    For i = 1 To auxdatos.ListItems.Count
        If lista.ListItems(lista.selectedItem.Index).Text = auxdatos.ListItems(i) And _
           lista.ListItems(lista.selectedItem.Index).SubItems(1) = auxdatos.ListItems(i).SubItems(1) And _
           lista.ListItems(lista.selectedItem.Index).SubItems(2) = auxdatos.ListItems(i).SubItems(2) Then
            datos.ListItems(CInt(auxdatos.ListItems(i).SubItems(4))).SubItems(1) = auxdatos.ListItems(i).SubItems(3)
        End If
    Next
    Set rs = Nothing
    Set rs_campos = Nothing
    Set ocampos = Nothing
    datos_Click

   On Error GoTo 0
   Exit Sub

cargar_campos_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure cargar_campos of Formulario frmCE_Resultados"
End Sub
Private Sub grabar_auxdatos()
    Dim i As Integer
   On Error GoTo grabar_auxdatos_Error

    For i = auxdatos.ListItems.Count To 1 Step -1
       If lista.ListItems(lista.selectedItem.Index).Text = auxdatos.ListItems(i) And _
          lista.ListItems(lista.selectedItem.Index).SubItems(1) = auxdatos.ListItems(i).SubItems(1) And _
          lista.ListItems(lista.selectedItem.Index).SubItems(2) = auxdatos.ListItems(i).SubItems(2) Then
           auxdatos.ListItems.Remove (i)
       End If
    Next
    For i = 1 To datos.ListItems.Count
       With auxdatos.ListItems.Add(, , lista.ListItems(lista.selectedItem.Index).Text) ' DESIGNACION
             .SubItems(1) = lista.ListItems(lista.selectedItem.Index).SubItems(1) ' PROBETA
             .SubItems(2) = lista.ListItems(lista.selectedItem.Index).SubItems(2) ' AREA
             .SubItems(3) = datos.ListItems(i).SubItems(1) ' VALOR
             .SubItems(4) = i ' LINEA
             .SubItems(5) = datos.ListItems(i).SubItems(3) ' CAMPO
             If datos.ListItems(i).bold = True Then
                .bold = True
                ' Si es solucion, la subimoslas determinaciones
                If UCase(lblestado.Caption) <> "DUPLICADA" Then
                    If datos.ListItems(i).SubItems(1) <> "" Then
                        lista.ListItems(lista.selectedItem.Index).SubItems(6) = datos.ListItems(i).SubItems(1)
                    End If
                End If
             Else
                If UCase(lblestado.Caption) = "DUPLICADA" Then
                    If datos.ListItems(i).Text = "Resultado (MEDIA)" Then
                        .SubItems(6) = "M"
                    End If
                    'M1371-I
                    If datos.ListItems(datos.ListItems.Count - 1).SubItems(1) <> "" Then
                        lista.ListItems(lista.selectedItem.Index).SubItems(6) = datos.ListItems(datos.ListItems.Count - 1).SubItems(1)
                    End If
                    'If datos.ListItems(i).Text = "Revisi�n de duplicados" Then
                    '    .SubItems(6) = "REV."
                    'End If
                    'If datos.ListItems(datos.ListItems.Count - 2).SubItems(1) <> "" Then
                    '    lista.ListItems(lista.selectedItem.Index).SubItems(6) = datos.ListItems(datos.ListItems.Count - 2).SubItems(1)
                    'End If
                End If
             End If
       End With
    Next

   On Error GoTo 0
   Exit Sub

grabar_auxdatos_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure grabar_auxdatos of Formulario frmCE_Resultados"
End Sub
Private Sub visualizar_duplicados()
        ' Si la muestra es duplicada, visualizar resultados
        Dim numero_resultados As Integer
        Dim i As Integer
        Dim res1 As String
        Dim res2 As String
   On Error GoTo visualizar_duplicados_Error

        numero_resultados = 0
        If UCase(lblestado.Caption) = "DUPLICADA" Then
            For i = 1 To datos.ListItems.Count
                If datos.ListItems(i).bold = True Then
                    If Trim(datos.ListItems(i).SubItems(1)) <> "" Then
                        numero_resultados = numero_resultados + 1
                        If Trim(res1) = "" Then
                            res1 = datos.ListItems(i).SubItems(1)
                        Else
                            res2 = datos.ListItems(i).SubItems(1)
                        End If
                    End If
                End If
            Next
        End If
        If numero_resultados = 2 And IsNumeric(res1) And IsNumeric(res2) Then ' Calcular media y diferencia
            Dim media As Single
            Dim dif As Single
            media = (CSng(res1) + CSng(res2)) / 2
            'M1371-I
            datos.ListItems(datos.ListItems.Count - 1).SubItems(1) = Format(CStr(media), "##0.00")
            'datos.ListItems(datos.ListItems.Count - 2).SubItems(1) = Format(CStr(media), "##0.00")
            'M1371-f
            grabar_auxdatos
            dif = Abs((CSng(res1) - CSng(res2)))
            'M1371-I
            datos.ListItems(datos.ListItems.Count).SubItems(1) = Format(CStr(dif), "#,##0.00")
            'datos.ListItems(datos.ListItems.Count - 1).SubItems(1) = Format(CStr(dif), "#,##0.00")
            'M1371-F
            grabar_auxdatos
        Else
            If res1 = "--" Or res2 = "--" Then
                'M1371-I
                'datos.ListItems(datos.ListItems.Count - 2).SubItems(1) = "--"
                'M1371-F
                datos.ListItems(datos.ListItems.Count - 1).SubItems(1) = "--"
                datos.ListItems(datos.ListItems.Count).SubItems(1) = "--"
            Else
                If UCase(lblestado.Caption) = "DUPLICADA" Then
                    'M1371-I
                    datos.ListItems(datos.ListItems.Count - 1).SubItems(1) = lista.ListItems(lista.selectedItem.Index).SubItems(6)
                    'datos.ListItems(datos.ListItems.Count - 1).SubItems(1) = lista.ListItems(lista.selectedItem.Index).SubItems(6)
                    'M1371-F
                End If
            End If
        End If

   On Error GoTo 0
   Exit Sub

visualizar_duplicados_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure visualizar_duplicados of Formulario frmCE_Resultados"
End Sub


Private Sub permisos()
   On Error GoTo permisos_Error

    If UCase(USUARIO.getUSUARIO) = "JULIO" Then
        txtformula.visible = True
        chkDuplicada.visible = True
    End If
    
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

   On Error GoTo 0
   Exit Sub

permisos_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure permisos of Formulario frmCE_Resultados"
    
End Sub
Private Sub almacenar_resultados_determinaciones()
    Dim i As Integer
    ' Almacenar Datos Determinaciones
    Dim oCe_RV As New clsCe_resultados_determinaciones
   On Error GoTo almacenar_resultados_determinaciones_Error
    If chkDuplicada.Value = Checked Then
        auxdatos.Sorted = True
        auxdatos.SortKey = 5
    End If

    For i = 1 To auxdatos.ListItems.Count
        If auxdatos.ListItems(i).SubItems(5) <> "" Then ' Para la media y diferencia de duplicados
            With oCe_RV
                .setMUESTRA_ID = PK_ID_MUESTRA
                .setDESIGNACION = auxdatos.ListItems(i).Text
                .setPROBETA = auxdatos.ListItems(i).SubItems(1)
                .setAREA = auxdatos.ListItems(i).SubItems(2)
                .setCAMPO_ID = auxdatos.ListItems(i).SubItems(5)
                .setVALOR_1 = " "
                If Trim(auxdatos.ListItems(i).SubItems(3)) <> "" Then
                    .setVALOR_1 = Replace(auxdatos.ListItems(i).SubItems(3), ",", ".")
                End If
                ' Valor duplicado
                .setVALOR_2 = " "
                If chkDuplicada.Value = Checked Then
                    i = i + 1
                    If Trim(auxdatos.ListItems(i).SubItems(3)) <> "" Then
                       .setVALOR_2 = Replace(auxdatos.ListItems(i).SubItems(3), ",", ".")
                    End If
                End If
                .Insertar
            End With
        End If
    Next
    ' Almacena en CE_resultados la Solucion
    Dim oCe_resultados As New clsCe_resultados
    With oCe_resultados
        For i = 1 To auxdatos.ListItems.Count
         If UCase(lblestado.Caption) = "DUPLICADA" Then
            If auxdatos.ListItems(i).SubItems(6) = "M" Then
                .setCONFORME = verificar_conforme(auxdatos.ListItems(i).SubItems(3))
                .setRESULTADO = Replace(auxdatos.ListItems(i).SubItems(3), ",", ".")
                .setFECHA = Format(Date, "dd/mm/yyyy")
                .Modificar_Resultado PK_ID_MUESTRA, auxdatos.ListItems(i).Text, auxdatos.ListItems(i).SubItems(1), auxdatos.ListItems(i).SubItems(2), False
            End If
         Else
            If auxdatos.ListItems(i).bold = True Then
                If IsNumeric(auxdatos.ListItems(i).SubItems(3)) Then
                    .setCONFORME = verificar_conforme(auxdatos.ListItems(i).SubItems(3))
                Else
                    .setCONFORME = 1
                End If
                .setRESULTADO = Replace(auxdatos.ListItems(i).SubItems(3), ",", ".")
                .setFECHA = Format(Date, "dd/mm/yyyy")
                .Modificar_Resultado PK_ID_MUESTRA, auxdatos.ListItems(i).Text, auxdatos.ListItems(i).SubItems(1), auxdatos.ListItems(i).SubItems(2), False
            End If
         End If
        Next
    End With
    Set oCe_RV = Nothing
    Set oCe_resultados = Nothing
   On Error GoTo 0
   Exit Sub

almacenar_resultados_determinaciones_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure almacenar_resultados_determinaciones of Formulario frmCE_Resultados"

End Sub
Private Function verificar_conforme(RESULTADO As Single) As Integer
   On Error GoTo verificar_conforme_Error

    verificar_conforme = 1
    If Trim(txtDatos(3)) <> "" And IsNumeric(txtDatos(3)) Then
        If CSng(Replace(txtDatos(3), ".", ",")) > CSng(Replace(RESULTADO, ".", ",")) Then
            verificar_conforme = 0
        End If
    End If
    If Trim(txtDatos(4)) <> "" And IsNumeric(txtDatos(4)) Then
        If CSng(Replace(txtDatos(4), ".", ",")) < CSng(Replace(RESULTADO, ".", ",")) Then
            verificar_conforme = 0
        End If
    End If

   On Error GoTo 0
   Exit Function

verificar_conforme_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure verificar_conforme of Formulario frmCE_Resultados"
End Function
Private Sub pasar_siguiente_campo()
   On Error GoTo pasar_siguiente_campo_Error

    If datos.ListItems.Count > datos.selectedItem.Index Then
        Set datos.selectedItem = datos.ListItems(datos.selectedItem.Index + 1)
        datos_Click
    Else
        If lista.ListItems.Count > lista.selectedItem.Index Then
            Set lista.selectedItem = lista.ListItems(lista.selectedItem.Index + 1)
            lista_Click
            datos_Click
        Else
            txtdato = ""
            txtValor = ""
            datos.SetFocus
        End If
    End If

   On Error GoTo 0
   Exit Sub

pasar_siguiente_campo_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure pasar_siguiente_campo of Formulario frmCE_Resultados"
End Sub
Private Sub cargar_equipos(MUESTRA As Long)
    Dim oCE As New clsCe_recepcion_equipos
    Dim rs As ADODB.Recordset
   On Error GoTo cargar_equipos_Error

    Set rs = oCE.Listado(MUESTRA)
    listaEquipos.ListItems.Clear
    If rs.RecordCount > 0 Then
        Do
            With listaEquipos.ListItems.Add(, , rs(0))
                .SubItems(1) = rs(1)
                .SubItems(2) = rs(2)
                .SubItems(3) = rs(5) ' VERIFICACION
                .SubItems(4) = rs(6) ' USOS MAX
                .SubItems(5) = rs(7) ' USOS
            End With
            If rs("EN_INFORME") = 1 Then
                listaEquipos.ListItems(listaEquipos.ListItems.Count).Checked = True
            Else
                listaEquipos.ListItems(listaEquipos.ListItems.Count).Checked = False
            End If
            rs.MoveNext
        Loop Until rs.EOF
    End If
    Set oCE = Nothing

   On Error GoTo 0
   Exit Sub

cargar_equipos_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure cargar_equipos of Formulario frmCE_Resultados"
End Sub

Private Sub almacenar_equipos()
    ' Insertar equipos en la tabla relacionada para el informe
    Dim i As Integer
    Dim oCE_Equipos As New clsCe_recepcion_equipos
   On Error GoTo almacenar_equipos_Error

    oCE_Equipos.Eliminar PK_ID_MUESTRA
    For i = 1 To listaEquipos.ListItems.Count
        With oCE_Equipos
            .setMUESTRA_ID = PK_ID_MUESTRA
            .setORDEN = i
            .setEQUIPO_ID = listaEquipos.ListItems(i).Text
            .setVERIFICACION_ID = listaEquipos.ListItems(i).SubItems(3)
'            .setEN_INFORME = Abs(listaEquipos.ListItems(i).Checked)
            .setEN_INFORME = Abs(listaEquipos.ListItems(i).Checked)
            .Insertar
        End With
    Next
    ' Usos de los equipos
    Dim oEU As New clsEq_usos
    oEU.Eliminar gmuestra, 0
    For i = 1 To listaEquipos.ListItems.Count
        With oEU
            .setMUESTRA_ID = PK_ID_MUESTRA
            .setDETERMINACION_ID = 0
            .setEQUIPO_ID = listaEquipos.ListItems(i).Text
            .setUSOS = listaEquipos.ListItems(i).SubItems(5)
'            If CInt(txtnumprobetas) = 1 Then
'                .setUSOS = CInt(txtnumprobetas)
'            Else
'                .setUSOS = CInt(txtnumprobetas) + 1
'            End If
            .Insertar
        End With
        ' Validar el n�mero de usos del equipo
        Dim oEquipo As New clsEquipos
        oEquipo.Carga listaEquipos.ListItems(i).Text
        If oEquipo.getNUMERO_USOS_MAXIMO <> 0 Then
            If oEquipo.getNUMERO_USOS_CONTADOR >= oEquipo.getNUMERO_USOS_MAXIMO Then
                MsgBox "ATENCION : El equipo N� " & oEquipo.getID_EQUIPO & " se ha usado " & oEquipo.getNUMERO_USOS_CONTADOR & " y su m�ximo es de " & oEquipo.getNUMERO_USOS_MAXIMO, vbCritical, App.Title
                Exit Sub
            End If
        End If
    Next
    Set oEU = Nothing

   On Error GoTo 0
   Exit Sub

almacenar_equipos_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure almacenar_equipos of Formulario frmCE_Resultados"
End Sub
