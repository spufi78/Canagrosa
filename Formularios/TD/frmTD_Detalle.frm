VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{CEA3B7F6-2847-4E5E-A551-DB7A62489D44}#46.0#0"; "miCombo.ocx"
Object = "{77DDF307-D82B-4757-8B3A-106EC9D75D4B}#8.0#0"; "tdbg8.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#13.2#0"; "Codejock.Controls.v13.2.1.ocx"
Begin VB.Form frmTD_Detalle 
   BackColor       =   &H80000004&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Tipos de Determinación"
   ClientHeight    =   10035
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   13200
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "frmTD_Detalle.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   10035
   ScaleWidth      =   13200
   StartUpPosition =   2  'CenterScreen
   Begin XtremeSuiteControls.TabControl TabControl 
      Height          =   4920
      Left            =   45
      TabIndex        =   24
      Top             =   4140
      Width           =   13110
      _Version        =   851970
      _ExtentX        =   23125
      _ExtentY        =   8678
      _StockProps     =   68
      Appearance      =   9
      Color           =   4
      PaintManager.BoldSelected=   -1  'True
      ItemCount       =   5
      Item(0).Caption =   "Datos Específicos"
      Item(0).ControlCount=   2
      Item(0).Control(0)=   "Frame3"
      Item(0).Control(1)=   "grid"
      Item(1).Caption =   "Equipos"
      Item(1).ControlCount=   1
      Item(1).Control(0)=   "Frame4"
      Item(2).Caption =   "Reactivos"
      Item(2).ControlCount=   1
      Item(2).Control(0)=   "Frame5"
      Item(3).Caption =   "Subcontratación"
      Item(3).ControlCount=   2
      Item(3).Control(0)=   "chkSubcontratable"
      Item(3).Control(1)=   "marcoSubcontratacion"
      Item(4).Caption =   "Tarifa"
      Item(4).ControlCount=   1
      Item(4).Control(0)=   "Frame2"
      Begin VB.Frame Frame2 
         Caption         =   "Datos Económicos"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   4485
         Left            =   -69955
         TabIndex        =   54
         Top             =   360
         Visible         =   0   'False
         Width           =   10815
         Begin VB.TextBox txtDatos 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
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
            Height          =   330
            Index           =   8
            Left            =   1080
            TabIndex        =   57
            Top             =   360
            Width           =   1575
         End
         Begin VB.CheckBox chkrevisarfactura 
            Caption         =   "Revisar factura"
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
            Height          =   240
            Left            =   90
            TabIndex        =   56
            Top             =   1215
            Width           =   1905
         End
         Begin VB.TextBox txttarifa 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
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
            Height          =   330
            Left            =   9405
            TabIndex        =   55
            Top             =   4050
            Width           =   1275
         End
         Begin MSComctlLib.ListView tarifas 
            Height          =   3810
            Left            =   5625
            TabIndex        =   58
            Top             =   180
            Width           =   5085
            _ExtentX        =   8969
            _ExtentY        =   6720
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
         Begin pryCombo.miCombo cmbtarifa 
            Height          =   375
            Left            =   1080
            TabIndex        =   59
            Top             =   765
            Width           =   4260
            _ExtentX        =   7514
            _ExtentY        =   661
         End
         Begin VB.Label lblCampos 
            AutoSize        =   -1  'True
            Caption         =   "Cód. Tarifa"
            Height          =   195
            Index           =   14
            Left            =   90
            TabIndex        =   62
            Top             =   855
            Width           =   780
         End
         Begin VB.Label lblCampos 
            AutoSize        =   -1  'True
            Caption         =   "Precio Base"
            Height          =   195
            Index           =   7
            Left            =   90
            TabIndex        =   61
            Top             =   420
            Width           =   855
         End
         Begin VB.Label lblCampos 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Precio Tarifa"
            Height          =   195
            Index           =   8
            Left            =   5895
            TabIndex        =   60
            Top             =   4140
            Width           =   3375
         End
      End
      Begin VB.Frame marcoSubcontratacion 
         Height          =   4020
         Left            =   -69910
         TabIndex        =   43
         Top             =   765
         Visible         =   0   'False
         Width           =   12840
         Begin VB.CommandButton cmdAnadirSubcontrata 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Añadir"
            Height          =   810
            Left            =   11745
            Picture         =   "frmTD_Detalle.frx":000C
            Style           =   1  'Graphical
            TabIndex        =   48
            Tag             =   "Añade campo o modifica el campo existente con el mismo nombre"
            Top             =   2970
            Width           =   915
         End
         Begin VB.CommandButton cmdEliminarSubcontrata 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Eliminar"
            Height          =   810
            Left            =   11745
            Picture         =   "frmTD_Detalle.frx":08D6
            Style           =   1  'Graphical
            TabIndex        =   47
            Tag             =   "Elimina el campo seleccionado"
            Top             =   225
            Width           =   915
         End
         Begin VB.TextBox txtDatos 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            Height          =   330
            Index           =   9
            Left            =   1080
            MaxLength       =   50
            TabIndex        =   46
            Top             =   3510
            Width           =   1215
         End
         Begin VB.TextBox txtDatos 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            Height          =   330
            Index           =   12
            Left            =   3105
            MaxLength       =   50
            TabIndex        =   45
            Top             =   3510
            Width           =   3555
         End
         Begin VB.TextBox txtDatos 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            Height          =   330
            Index           =   19
            Left            =   7650
            MaxLength       =   50
            TabIndex        =   44
            Top             =   3510
            Width           =   1305
         End
         Begin MSComctlLib.ListView listaSubcontratas 
            Height          =   2790
            Left            =   90
            TabIndex        =   49
            Top             =   180
            Width           =   11430
            _ExtentX        =   20161
            _ExtentY        =   4921
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
         Begin pryCombo.miCombo cmbSubcontratas 
            Height          =   330
            Left            =   1080
            TabIndex        =   50
            Top             =   3060
            Width           =   10410
            _ExtentX        =   18362
            _ExtentY        =   582
         End
         Begin VB.Label Label4 
            Caption         =   "Subcontrata"
            Height          =   285
            Left            =   90
            TabIndex        =   64
            Top             =   3105
            Width           =   915
         End
         Begin VB.Label Label1 
            Caption         =   "Valor Ref.:"
            Height          =   285
            Left            =   135
            TabIndex        =   53
            Top             =   3555
            Width           =   915
         End
         Begin VB.Label Label2 
            Caption         =   "Norma:"
            Height          =   285
            Left            =   2430
            TabIndex        =   52
            Top             =   3555
            Width           =   555
         End
         Begin VB.Label Label3 
            Caption         =   "Precio (€):"
            ForeColor       =   &H00000000&
            Height          =   285
            Left            =   6795
            TabIndex        =   51
            Top             =   3555
            Width           =   735
         End
      End
      Begin VB.CheckBox chkSubcontratable 
         Caption         =   "Tipo de Determinación Subcontratable"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   -69865
         TabIndex        =   42
         Top             =   450
         Visible         =   0   'False
         Width           =   3660
      End
      Begin VB.Frame Frame5 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   4350
         Left            =   -69910
         TabIndex        =   34
         Top             =   450
         Visible         =   0   'False
         Width           =   12840
         Begin VB.CommandButton cmdEliminarReactivo 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Eliminar"
            Height          =   750
            Left            =   11700
            Picture         =   "frmTD_Detalle.frx":11A0
            Style           =   1  'Graphical
            TabIndex        =   36
            Tag             =   "Elimina el campo seleccionado"
            Top             =   225
            Width           =   915
         End
         Begin VB.CommandButton cmdAnadirReactivo 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Añadir"
            Height          =   765
            Left            =   11745
            Picture         =   "frmTD_Detalle.frx":1A6A
            Style           =   1  'Graphical
            TabIndex        =   35
            Tag             =   "Añade campo o modifica el campo existente con el mismo nombre"
            Top             =   3420
            Width           =   915
         End
         Begin MSComctlLib.ListView listaReactivos 
            Height          =   3120
            Left            =   135
            TabIndex        =   37
            Top             =   270
            Width           =   11340
            _ExtentX        =   20003
            _ExtentY        =   5503
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
            Left            =   825
            TabIndex        =   38
            Top             =   3540
            Width           =   10695
            _ExtentX        =   18865
            _ExtentY        =   582
         End
         Begin pryCombo.miCombo cmbReactivosInternos 
            Height          =   330
            Left            =   825
            TabIndex        =   39
            Top             =   3870
            Width           =   10695
            _ExtentX        =   18865
            _ExtentY        =   582
         End
         Begin VB.Label lblCampos 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C0C0&
            Caption         =   "Interno"
            Height          =   195
            Index           =   25
            Left            =   150
            TabIndex        =   41
            Top             =   3915
            Width           =   495
         End
         Begin VB.Label lblCampos 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C0C0&
            Caption         =   "Externos"
            Height          =   195
            Index           =   26
            Left            =   150
            TabIndex        =   40
            Top             =   3570
            Width           =   615
         End
      End
      Begin VB.Frame Frame4 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   4305
         Left            =   -69910
         TabIndex        =   29
         Top             =   540
         Visible         =   0   'False
         Width           =   12885
         Begin VB.CommandButton cmdEliminarEquipo 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Eliminar"
            Height          =   810
            Left            =   11610
            Picture         =   "frmTD_Detalle.frx":2334
            Style           =   1  'Graphical
            TabIndex        =   31
            Tag             =   "Elimina el campo seleccionado"
            Top             =   270
            Width           =   915
         End
         Begin VB.CommandButton cmdAnadirEquipo 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Añadir"
            Height          =   810
            Left            =   11610
            Picture         =   "frmTD_Detalle.frx":2BFE
            Style           =   1  'Graphical
            TabIndex        =   30
            Tag             =   "Añade campo o modifica el campo existente con el mismo nombre"
            Top             =   3330
            Width           =   915
         End
         Begin MSComctlLib.ListView listaEquipos 
            Height          =   3465
            Left            =   135
            TabIndex        =   32
            Top             =   270
            Width           =   11205
            _ExtentX        =   19764
            _ExtentY        =   6112
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
         Begin pryCombo.miCombo cmbEquipos 
            Height          =   330
            Left            =   150
            TabIndex        =   33
            Top             =   3825
            Width           =   11175
            _ExtentX        =   19711
            _ExtentY        =   582
         End
      End
      Begin VB.Frame Frame3 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1110
         Left            =   90
         TabIndex        =   25
         Top             =   495
         Width           =   12930
         Begin VB.CheckBox chkparticulas 
            Caption         =   "Contaminación de partículas"
            Height          =   285
            Left            =   90
            TabIndex        =   27
            Top             =   225
            Width           =   2355
         End
         Begin VB.TextBox txtDatos 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
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
            Height          =   330
            Index           =   15
            Left            =   810
            TabIndex        =   26
            Top             =   630
            Width           =   12015
         End
         Begin VB.Label lblCampos 
            AutoSize        =   -1  'True
            BackColor       =   &H80000004&
            Caption         =   "Método"
            Height          =   195
            Index           =   22
            Left            =   90
            TabIndex        =   28
            Top             =   675
            Width           =   540
         End
      End
      Begin TrueDBGrid80.TDBGrid grid 
         Height          =   3225
         Left            =   90
         TabIndex        =   63
         Top             =   1620
         Width           =   12930
         _ExtentX        =   22807
         _ExtentY        =   5689
         _LayoutType     =   4
         _RowHeight      =   -2147483647
         _WasPersistedAsPixels=   0
         Columns(0)._VlistStyle=   0
         Columns(0)._MaxComboItems=   5
         Columns(0).Caption=   "ORDEN"
         Columns(0).DataField=   ""
         Columns(0).NumberFormat=   "General Number"
         Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(1)._VlistStyle=   0
         Columns(1)._MaxComboItems=   5
         Columns(1).Caption=   "Rango"
         Columns(1).DataField=   ""
         Columns(1).DropDown=   "tMetodos"
         Columns(1).DropDown.vt=   8
         Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(2)._VlistStyle=   0
         Columns(2)._MaxComboItems=   5
         Columns(2).Caption=   "Incertidumbre"
         Columns(2).DataField=   ""
         Columns(2).NumberFormat=   "Fixed"
         Columns(2).ExternalEditor=   "TDBDate1"
         Columns(2).ExternalEditor.vt=   8
         Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(3)._VlistStyle=   0
         Columns(3)._MaxComboItems=   5
         Columns(3).Caption=   "Dif.Duplicados"
         Columns(3).DataField=   ""
         Columns(3).NumberFormat=   "General Number"
         Columns(3).DropDown=   "tUnidades"
         Columns(3).DropDown.vt=   8
         Columns(3).ExternalEditor=   "TDBDate1"
         Columns(3).ExternalEditor.vt=   8
         Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(4)._VlistStyle=   0
         Columns(4)._MaxComboItems=   5
         Columns(4).Caption=   "% Dif.Duplicados"
         Columns(4).DataField=   ""
         Columns(4).NumberFormat=   "Fixed"
         Columns(4)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(5)._VlistStyle=   0
         Columns(5)._MaxComboItems=   5
         Columns(5).Caption=   "% Aviso Rango"
         Columns(5).DataField=   ""
         Columns(5).NumberFormat=   "General Number"
         Columns(5)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(6)._VlistStyle=   0
         Columns(6)._MaxComboItems=   5
         Columns(6).Caption=   "% Dif. Histórico"
         Columns(6).DataField=   ""
         Columns(6).NumberFormat=   "General Number"
         Columns(6)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(7)._VlistStyle=   0
         Columns(7)._MaxComboItems=   5
         Columns(7).Caption=   "Coef. Variación"
         Columns(7).DataField=   ""
         Columns(7).NumberFormat=   "General Number"
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
         Splits(0)._ColumnProps(1)=   "Column(0).Width=1217"
         Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
         Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=1111"
         Splits(0)._ColumnProps(4)=   "Column(0)._EditAlways=0"
         Splits(0)._ColumnProps(5)=   "Column(0).Visible=0"
         Splits(0)._ColumnProps(6)=   "Column(0).Order=1"
         Splits(0)._ColumnProps(7)=   "Column(1).Width=5133"
         Splits(0)._ColumnProps(8)=   "Column(1).DividerColor=0"
         Splits(0)._ColumnProps(9)=   "Column(1)._WidthInPix=5027"
         Splits(0)._ColumnProps(10)=   "Column(1)._EditAlways=0"
         Splits(0)._ColumnProps(11)=   "Column(1)._ColStyle=1"
         Splits(0)._ColumnProps(12)=   "Column(1).Order=2"
         Splits(0)._ColumnProps(13)=   "Column(2).Width=2170"
         Splits(0)._ColumnProps(14)=   "Column(2).DividerColor=0"
         Splits(0)._ColumnProps(15)=   "Column(2)._WidthInPix=2064"
         Splits(0)._ColumnProps(16)=   "Column(2)._EditAlways=0"
         Splits(0)._ColumnProps(17)=   "Column(2)._ColStyle=1"
         Splits(0)._ColumnProps(18)=   "Column(2).Order=3"
         Splits(0)._ColumnProps(19)=   "Column(3).Width=5794"
         Splits(0)._ColumnProps(20)=   "Column(3).DividerColor=0"
         Splits(0)._ColumnProps(21)=   "Column(3)._WidthInPix=5689"
         Splits(0)._ColumnProps(22)=   "Column(3)._EditAlways=0"
         Splits(0)._ColumnProps(23)=   "Column(3)._ColStyle=1"
         Splits(0)._ColumnProps(24)=   "Column(3).Order=4"
         Splits(0)._ColumnProps(25)=   "Column(3).AutoDropDown=1"
         Splits(0)._ColumnProps(26)=   "Column(4).Width=2381"
         Splits(0)._ColumnProps(27)=   "Column(4).DividerColor=0"
         Splits(0)._ColumnProps(28)=   "Column(4)._WidthInPix=2275"
         Splits(0)._ColumnProps(29)=   "Column(4)._EditAlways=0"
         Splits(0)._ColumnProps(30)=   "Column(4)._ColStyle=1"
         Splits(0)._ColumnProps(31)=   "Column(4).Order=5"
         Splits(0)._ColumnProps(32)=   "Column(5).Width=2434"
         Splits(0)._ColumnProps(33)=   "Column(5).DividerColor=0"
         Splits(0)._ColumnProps(34)=   "Column(5)._WidthInPix=2328"
         Splits(0)._ColumnProps(35)=   "Column(5)._EditAlways=0"
         Splits(0)._ColumnProps(36)=   "Column(5)._ColStyle=1"
         Splits(0)._ColumnProps(37)=   "Column(5).Order=6"
         Splits(0)._ColumnProps(38)=   "Column(6).Width=2381"
         Splits(0)._ColumnProps(39)=   "Column(6).DividerColor=0"
         Splits(0)._ColumnProps(40)=   "Column(6)._WidthInPix=2275"
         Splits(0)._ColumnProps(41)=   "Column(6)._EditAlways=0"
         Splits(0)._ColumnProps(42)=   "Column(6)._ColStyle=1"
         Splits(0)._ColumnProps(43)=   "Column(6).Order=7"
         Splits(0)._ColumnProps(44)=   "Column(7).Width=1323"
         Splits(0)._ColumnProps(45)=   "Column(7).DividerColor=0"
         Splits(0)._ColumnProps(46)=   "Column(7)._WidthInPix=1217"
         Splits(0)._ColumnProps(47)=   "Column(7)._EditAlways=0"
         Splits(0)._ColumnProps(48)=   "Column(7)._ColStyle=1"
         Splits(0)._ColumnProps(49)=   "Column(7).Order=8"
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
         _StyleDefs(36)  =   "Splits(0).Columns(0).Style:id=62,.parent=11"
         _StyleDefs(37)  =   "Splits(0).Columns(0).HeadingStyle:id=59,.parent=12"
         _StyleDefs(38)  =   "Splits(0).Columns(0).FooterStyle:id=60,.parent=13"
         _StyleDefs(39)  =   "Splits(0).Columns(0).EditorStyle:id=61,.parent=15"
         _StyleDefs(40)  =   "Splits(0).Columns(1).Style:id=32,.parent=11,.alignment=2,.locked=0"
         _StyleDefs(41)  =   "Splits(0).Columns(1).HeadingStyle:id=29,.parent=12"
         _StyleDefs(42)  =   "Splits(0).Columns(1).FooterStyle:id=30,.parent=13"
         _StyleDefs(43)  =   "Splits(0).Columns(1).EditorStyle:id=31,.parent=15"
         _StyleDefs(44)  =   "Splits(0).Columns(2).Style:id=54,.parent=11,.alignment=2"
         _StyleDefs(45)  =   "Splits(0).Columns(2).HeadingStyle:id=51,.parent=12"
         _StyleDefs(46)  =   "Splits(0).Columns(2).FooterStyle:id=52,.parent=13"
         _StyleDefs(47)  =   "Splits(0).Columns(2).EditorStyle:id=53,.parent=15"
         _StyleDefs(48)  =   "Splits(0).Columns(3).Style:id=28,.parent=11,.alignment=2"
         _StyleDefs(49)  =   "Splits(0).Columns(3).HeadingStyle:id=25,.parent=12"
         _StyleDefs(50)  =   "Splits(0).Columns(3).FooterStyle:id=26,.parent=13"
         _StyleDefs(51)  =   "Splits(0).Columns(3).EditorStyle:id=27,.parent=15"
         _StyleDefs(52)  =   "Splits(0).Columns(4).Style:id=66,.parent=11,.alignment=2"
         _StyleDefs(53)  =   "Splits(0).Columns(4).HeadingStyle:id=63,.parent=12"
         _StyleDefs(54)  =   "Splits(0).Columns(4).FooterStyle:id=64,.parent=13"
         _StyleDefs(55)  =   "Splits(0).Columns(4).EditorStyle:id=65,.parent=15"
         _StyleDefs(56)  =   "Splits(0).Columns(5).Style:id=36,.parent=11,.alignment=2"
         _StyleDefs(57)  =   "Splits(0).Columns(5).HeadingStyle:id=33,.parent=12"
         _StyleDefs(58)  =   "Splits(0).Columns(5).FooterStyle:id=34,.parent=13"
         _StyleDefs(59)  =   "Splits(0).Columns(5).EditorStyle:id=35,.parent=15"
         _StyleDefs(60)  =   "Splits(0).Columns(6).Style:id=24,.parent=11,.alignment=2"
         _StyleDefs(61)  =   "Splits(0).Columns(6).HeadingStyle:id=21,.parent=12"
         _StyleDefs(62)  =   "Splits(0).Columns(6).FooterStyle:id=22,.parent=13"
         _StyleDefs(63)  =   "Splits(0).Columns(6).EditorStyle:id=23,.parent=15"
         _StyleDefs(64)  =   "Splits(0).Columns(7).Style:id=58,.parent=11,.alignment=2"
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
         _StyleDefs(80)  =   ":id=40,.parent=37,.bgcolor=&H8080FF&,.fgcolor=&H0&,.bold=-1,.fontsize=975"
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
   End
   Begin VB.CommandButton cmdHistorialCambios 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Historial Cambios"
      Height          =   870
      Left            =   8550
      Style           =   1  'Graphical
      TabIndex        =   23
      Top             =   9090
      Width           =   1365
   End
   Begin VB.CommandButton cmdQuien 
      BackColor       =   &H00E0E0E0&
      Caption         =   "¿Donde?"
      Height          =   870
      Left            =   9945
      Style           =   1  'Graphical
      TabIndex        =   20
      Top             =   9090
      Width           =   1050
   End
   Begin VB.CommandButton cmdok 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Aceptar"
      Height          =   870
      Left            =   11025
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   9090
      Width           =   1050
   End
   Begin VB.CommandButton cmdcancel 
      BackColor       =   &H00E0E0E0&
      Caption         =   "ESC-Salir"
      Height          =   870
      Left            =   12105
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   9090
      Width           =   1050
   End
   Begin VB.Frame Frame1 
      Height          =   3390
      Left            =   45
      TabIndex        =   9
      Top             =   675
      Width           =   13110
      Begin pryCombo.miCombo cmbFormula 
         Height          =   330
         Left            =   1350
         TabIndex        =   3
         Top             =   1425
         Width           =   11670
         _ExtentX        =   20585
         _ExtentY        =   582
      End
      Begin VB.TextBox txtDatos 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   330
         Index           =   7
         Left            =   1350
         TabIndex        =   1
         Top             =   540
         Width           =   11670
      End
      Begin VB.TextBox txtDatos 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Height          =   330
         Index           =   3
         Left            =   6975
         TabIndex        =   8
         Top             =   2970
         Width           =   6045
      End
      Begin VB.TextBox txtDatos 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Height          =   330
         Index           =   2
         Left            =   1350
         MaxLength       =   50
         TabIndex        =   7
         Top             =   2970
         Width           =   4785
      End
      Begin VB.TextBox txtDatos 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
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
         Index           =   0
         Left            =   1350
         TabIndex        =   0
         Top             =   180
         Width           =   11670
      End
      Begin VB.TextBox txtDatos 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   465
         Index           =   1
         Left            =   1350
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   2
         Top             =   915
         Width           =   11670
      End
      Begin pryCombo.miCombo cmbFamilia 
         Height          =   330
         Left            =   1350
         TabIndex        =   4
         Top             =   1785
         Width           =   11670
         _ExtentX        =   20585
         _ExtentY        =   582
      End
      Begin pryCombo.miCombo cmbPNT 
         Height          =   330
         Left            =   1350
         TabIndex        =   5
         Top             =   2160
         Width           =   11670
         _ExtentX        =   20585
         _ExtentY        =   582
      End
      Begin pryCombo.miCombo cmbNorma 
         Height          =   330
         Left            =   1350
         TabIndex        =   6
         Top             =   2565
         Width           =   11670
         _ExtentX        =   20585
         _ExtentY        =   582
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H80000004&
         Caption         =   "Norma Vinculada"
         Height          =   195
         Index           =   5
         Left            =   90
         TabIndex        =   65
         Top             =   2610
         Width           =   1215
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H80000004&
         Caption         =   "PNT Informe"
         Height          =   195
         Index           =   12
         Left            =   90
         TabIndex        =   22
         Top             =   3060
         Width           =   900
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H80000004&
         Caption         =   "Familia"
         Height          =   195
         Index           =   16
         Left            =   90
         TabIndex        =   21
         Top             =   1830
         Width           =   480
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H80000004&
         Caption         =   "Inglés"
         Height          =   195
         Index           =   6
         Left            =   90
         TabIndex        =   17
         Top             =   585
         Width           =   435
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H80000004&
         Caption         =   "Norma"
         Height          =   195
         Index           =   4
         Left            =   6345
         TabIndex        =   16
         Top             =   3015
         Width           =   465
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H80000004&
         Caption         =   "PNT Vinculado"
         Height          =   195
         Index           =   1
         Left            =   90
         TabIndex        =   15
         Top             =   2205
         Width           =   1080
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H80000004&
         Caption         =   "Formula"
         Height          =   195
         Index           =   0
         Left            =   90
         TabIndex        =   14
         Top             =   1455
         Width           =   555
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H80000004&
         Caption         =   "Nombre"
         Height          =   195
         Index           =   2
         Left            =   90
         TabIndex        =   13
         Top             =   210
         Width           =   585
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H80000004&
         Caption         =   "Descripción"
         Height          =   195
         Index           =   3
         Left            =   90
         TabIndex        =   12
         Top             =   1035
         Width           =   840
      End
   End
   Begin VB.Image imagen 
      Height          =   480
      Left            =   12555
      Picture         =   "frmTD_Detalle.frx":34C8
      Top             =   90
      Width           =   480
   End
   Begin VB.Label lbltitulo 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "Detalle de Tipos de determinación"
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
      TabIndex        =   19
      Top             =   75
      Width           =   3615
   End
   Begin VB.Label lblsubtitulo 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "Detalle de los tipos de determinaciones"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   90
      TabIndex        =   18
      Top             =   375
      Width           =   2760
   End
   Begin VB.Shape fondo 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   690
      Left            =   -90
      Top             =   0
      Width           =   13415
   End
End
Attribute VB_Name = "frmTD_Detalle"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public PK As Long
Private tarifa_modificada As Boolean

Dim gridTabla As New XArrayDB
Const filas As Integer = 100
Const Col As Integer = 8
Private Enum COLS
    C_ORDEN = 0
    C_RANGO = 1
    C_INCERTIDUMBRE = 2
    C_DIF_DUPLICADOS_NUMERICA = 3
    C_DIF_DUPLICADOS = 4
    C_DIF_AVISO = 5
    C_DIF_HISTORICO = 6
    C_C_VARIACION = 7
End Enum
Private Sub cabecera()
    With tarifas.ColumnHeaders
        .Add , , "ID", 1, lvwColumnLeft
        .Add , , "Tarifa", 3500, lvwColumnLeft
        .Add , , "Precio", 1275, lvwColumnRight
    End With
    With listaEquipos.ColumnHeaders
        .Add , , "NºEquipo", 900, lvwColumnLeft
        .Add , , "Nombre", 7000, lvwColumnLeft
        .Add , , "NºSerie", 2600, lvwColumnCenter
    End With
    With listaReactivos.ColumnHeaders
        .Add , , "Número", 800, lvwColumnLeft
        .Add , , "Reactivo", 7600, lvwColumnLeft
        .Add , , "Caducidad", 2500, lvwColumnCenter
        .Add , , "TIPO", 0, lvwColumnCenter
    End With
    With listaSubcontratas.ColumnHeaders
        .Add , , "ID", 1, lvwColumnLeft
        .Add , , "Proveedor", 5500, lvwColumnCenter
        .Add , , "Valor Ref.", 1100, lvwColumnCenter
        .Add , , "Norma", 2800, lvwColumnCenter
        .Add , , "Precio (€)", 1200, lvwColumnCenter
    End With
End Sub


Private Sub cmdAnadirSubcontrata_Click()
    If cmbSubcontratas.getPK_SALIDA <> 0 Then
        If txtDatos(9).Text = "" Then
           MsgBox "Introduzca el Valor de Referencia", vbInformation, App.Title
           txtDatos(9).SetFocus
           Exit Sub
        End If
        
        If txtDatos(12).Text = "" Then
           MsgBox "Introduzca la Normativa Aplicable", vbInformation, App.Title
           txtDatos(12).SetFocus
           Exit Sub
        End If
        
        Dim oProveedor As New clsProveedor
        oProveedor.Carga cmbSubcontratas.getPK_SALIDA
        ' Verificar si existe el formador
        Dim z As Integer
        For z = 1 To listaSubcontratas.ListItems.Count
             If CLng(listaSubcontratas.ListItems(z).Text) = CLng(cmbSubcontratas.getPK_SALIDA) Then
                 listaSubcontratas.ListItems(z).SubItems(2) = Trim(txtDatos(9))
                 listaSubcontratas.ListItems(z).SubItems(3) = Trim(txtDatos(12))
                 listaSubcontratas.ListItems(z).SubItems(4) = moneda(Trim(txtDatos(19)))
                 txtDatos(9).Text = ""
                 txtDatos(12).Text = ""
                 txtDatos(19).Text = ""
                 Exit Sub
             End If
        Next z
        
        With listaSubcontratas.ListItems.Add(, , cmbSubcontratas.getPK_SALIDA)
            .SubItems(1) = oProveedor.getNOMBRE
            .SubItems(2) = Trim(txtDatos(9))
            .SubItems(3) = Trim(txtDatos(12))
            .SubItems(4) = Trim(txtDatos(19))
        End With
        
        listaSubcontratas.ListItems(listaSubcontratas.ListItems.Count).EnsureVisible
        cmbSubcontratas.limpiar
        txtDatos(9).Text = ""
        txtDatos(12).Text = ""
        txtDatos(19).Text = ""
    End If
End Sub

'M0927-I
Private Sub cmdEliminarSubcontrata_Click()
    If listaSubcontratas.ListItems.Count > 0 Then
        listaSubcontratas.ListItems.Remove listaSubcontratas.selectedItem.Index
    End If
End Sub
'M0927-F

Private Sub cmdHistorialCambios_Click()
    If PK <> 0 Then
        frmHistorialCambios.PK_TIPO = HC_TIPOS.HC_TIPO_DETERMINACION
        frmHistorialCambios.PK_ID = PK
        frmHistorialCambios.PK_TITULO = "Tipo Determinación " & txtDatos(0)
        frmHistorialCambios.Show 1
    End If
End Sub
Private Sub cmbPNT_change()
    If cmbPNT.getPK_SALIDA <> 0 Then
        Dim oCA_Documento As New clsCa_documentos
        If oCA_Documento.Carga(cmbPNT.getPK_SALIDA) Then
            txtDatos(2) = oCA_Documento.getCODIGO
        End If
    End If
End Sub

Private Sub cmdAnadirEquipo_Click()
    If cmbEquipos.getPK_SALIDA <> 0 Then
        Dim oEquipo As New clsEquipos
        oEquipo.Carga cmbEquipos.getPK_SALIDA
        With listaEquipos.ListItems.Add(, , oEquipo.getID_EQUIPO)
            .SubItems(1) = oEquipo.getNOMBRE
            .SubItems(2) = oEquipo.getSERIE
        End With
        listaEquipos.ListItems(listaEquipos.ListItems.Count).EnsureVisible
        cmbEquipos.limpiar
    End If
End Sub

Private Sub cmdAnadirReactivo_Click()
    ' Externo (E)
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
    ' Interno (I)
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
End Sub

Private Sub cmdEliminarEquipo_Click()
    If listaEquipos.ListItems.Count > 0 Then
        listaEquipos.ListItems.Remove listaEquipos.selectedItem.Index
    End If
End Sub

Private Sub cmdEliminarReactivo_Click()
    If listaReactivos.ListItems.Count > 0 Then
        listaReactivos.ListItems.Remove listaReactivos.selectedItem.Index
        cmbReactivos.limpiar
        cmbReactivosInternos.limpiar
    End If
End Sub

Private Sub chkSubcontratable_Click()
    If chkSubcontratable.Value = Checked Then
        marcoSubcontratacion.Enabled = True
    Else
        marcoSubcontratacion.Enabled = False
    End If
End Sub
Private Sub cmdQuien_Click()
    If PK <> 0 Then
        frmTD_Donde.PK = PK
        frmTD_Donde.Show 1
    End If
End Sub

Private Sub anadir_precio()
    If tarifas.ListItems.Count > 0 Then
        If txttarifa.Text = "" Then
            MsgBox "Introduzca el precio correctamente.", vbInformation, App.Title
            txttarifa.SetFocus
        Else
            If moneda(tarifas.ListItems(tarifas.selectedItem.Index).SubItems(2)) <> moneda(txttarifa) Then
                tarifa_modificada = True
            End If
            tarifas.ListItems(tarifas.selectedItem.Index).SubItems(2) = moneda(txttarifa)
            txttarifa = ""
            If tarifas.ListItems.Count > tarifas.selectedItem.Index Then
                Set tarifas.selectedItem = tarifas.ListItems(tarifas.selectedItem.Index + 1)
                tarifas.SetFocus
                tarifas_Click
            End If
        End If
    End If
End Sub
Private Sub cmdcancel_Click()
    Unload Me
End Sub

Private Sub cmdok_Click()
   On Error GoTo cmdok_Click_Error

    If validar = True Then
      Dim oTD As New clsTipos_determinacion
      'M0927-I
      Dim oTDContratas As New clsTipos_determinacion_contratas
      'M0927-F
      Dim DETERMINACION As Long
      With oTD
            .setNOMBRE = txtDatos(0)
            .setNOMBRE_INGLES = txtDatos(7)
            .setDESCRIPCION = txtDatos(1)
            .setFORMULA_ID = cmbFormula.getPK_SALIDA
            .setFAMILIA_ID = cmbFamilia.getPK_SALIDA
            .setPNT_VINCULADO = cmbPNT.getPK_SALIDA
            .setPNT = txtDatos(2)
            .setPROC_REF_EADS = txtDatos(3)
            If txtDatos(8) <> "" Then
                .setPRECIO = moneda_bd(txtDatos(8))
            Else
                .setPRECIO = moneda_bd("0")
            End If
            .setPARTICULAS = chkparticulas.Value
            .setREVISAR_FACTURA = chkrevisarfactura.Value
            .setTARIFA_CODIGO_ID = cmbtarifa.getPK_SALIDA
            .setES_SUBCONTRATABLE = chkSubcontratable.Value
            .setNORMA_ID = cmbnorma.getPK_SALIDA
            
            oTDContratas.Eliminar_Tipo PK
            If listaSubcontratas.ListItems.Count > 0 Then
                Dim j As Integer
                For j = 1 To listaSubcontratas.ListItems.Count
                    oTDContratas.CrearID
                    oTDContratas.setCONTRATA_ID = CLng(listaSubcontratas.ListItems(j).Text)
                    oTDContratas.setNORMATIVA_APLICABLE = Trim(listaSubcontratas.ListItems(j).SubItems(3))
                    oTDContratas.setTIPO_DETERMINACION_ID = PK
                    oTDContratas.setVALOR_REFERENCIA = Trim(listaSubcontratas.ListItems(j).SubItems(2))
                    oTDContratas.setPRECIO = moneda_bd(listaSubcontratas.ListItems(j).SubItems(4))
                    oTDContratas.Insertar
                Next j
            End If
            .setMETODO = txtDatos(15)
      End With
      Dim ohc As New clsHistorial_cambios
      If PK = 0 Then
        If MsgBox("Va a introducir un nuevo tipo de determinación. ¿Ha revisado que todos los datos son conformes?", vbQuestion + vbYesNo, App.Title) = vbYes Then
            DETERMINACION = oTD.Insertar
            If DETERMINACION > 0 Then
                With ohc
                    .setTIPO = HC_TIPOS.HC_TIPO_DETERMINACION
                    .setIDENTIFICADOR = DETERMINACION
                    .setIDENTIFICADOR_TEXTO = txtDatos(0)
                    .setUSUARIO_ID = USUARIO.getID_EMPLEADO
                    .setMOTIVO = HC_CREACION
                    .Insertar
                End With
            End If
        Else
            Exit Sub
        End If
      Else
        If MsgBox("Va a modificar el tipo de determinación. ¿Ha revisado que todos los datos son conformes?", vbQuestion + vbYesNo, App.Title) = vbYes Then
            frmMotivo.lbltitulo = "Indique detalladamente el motivo de modificación del tipo de determinación."
            frmMotivo.Show 1
            If Trim(MOTIVO) = "" Then
                MsgBox "Para modificar los datos es necesario introducir el motivo de la modificación.", vbInformation, App.Title
                Exit Sub
            End If
            If oTD.Modificar(PK) = True Then
                DETERMINACION = PK
                With ohc
                    .setTIPO = HC_TIPOS.HC_TIPO_DETERMINACION
                    .setIDENTIFICADOR = PK
                    .setUSUARIO_ID = USUARIO.getID_EMPLEADO
                    .setIDENTIFICADOR_TEXTO = txtDatos(0)
                    .setMOTIVO = Trim(MOTIVO)
                    .Insertar
                End With
            End If
        Else
            Exit Sub
        End If
      End If
      Set ohc = Nothing
      ' Tarifas
      Dim i As Integer
      Me.MousePointer = 11

      ' Enviar correo si se modifica la tarifa
'      If tarifa_modificada = True Then
'            Dim oParametro As New clsParametros
'            oParametro.Carga PARAM_USUARIO_VIGILADO, ""
'
'            If USUARIO.getID_EMPLEADO = oParametro.getVALOR Then
'                Dim asunto As String
'                Dim DETALLE As String
'                asunto = "El usuario " & USUARIO.getUSUARIO & " ha modificado la tarifa de un tipo determinación."
'
'                DETALLE = "" & vbNewLine
'                DETALLE = DETALLE & " Fecha : " & Format(Date, "dd-mm-yyyy") & vbNewLine
'                DETALLE = DETALLE & " Hora  : " & Time & vbNewLine & vbNewLine
'                DETALLE = DETALLE & " Tipo Determinacion : " & txtDatos(0) & vbNewLine & vbNewLine
'
'                DETALLE = DETALLE & " Cambios en Tarifa " & vbNewLine
'                DETALLE = DETALLE & " ----------------- " & vbNewLine
'
'                Dim CO As String
'                Dim rs2 As ADODB.RecordSet
'                CO = "SELECT A.ID_TARIFA, A.NOMBRE, B.PRECIO " & _
'                     "  FROM TARIFAS A LEFT JOIN TARIFAS_PRECIOS B ON A.ID_TARIFA = B.TARIFA_ID  AND B.TIPO_DETERMINACION_ID = " & DETERMINACION & _
'                     " where A.EN_VIGOR = 1 "
'                Set rs2 = datos_bd(CO)
'                Dim PRECIO As String
'                Dim precio_ant As String
'                If rs2.RecordCount > 0 Then
'                    Do
'                            For i = 1 To tarifas.ListItems.Count
'                              If tarifas.ListItems(i).Text = rs2(0) Then
'                                If IsNull(rs2(2)) Then
'                                    precio_ant = moneda("0")
'                                Else
'                                    precio_ant = moneda(rs2(2))
'                                End If
'                                If Trim(tarifas.ListItems(i).SubItems(2)) = "" Then
'                                    PRECIO = moneda("0")
'                                Else
'                                    PRECIO = moneda(tarifas.ListItems(i).SubItems(2))
'                                End If
'                                If PRECIO <> precio_ant Then
'                                    DETALLE = DETALLE & tarifas.ListItems(i).SubItems(1) & " : " & precio_ant & " -> " & PRECIO & vbNewLine
'                                End If
'                              End If
'                            Next
'                        rs2.MoveNext
'                    Loop Until rs2.EOF
'                End If
'                Dim CO As String
'                Dim rs2 As ADODB.RecordSet
'                CO = "SELECT A.NOMBRE, B.PRECIO " & _
'                     "  FROM TARIFAS A, TARIFAS_PRECIOS B " & _
'                     " where a.ID_TARIFA = b.TARIFA_ID AND A.EN_VIGOR = 1 " & _
'                     "   AND B.TIPO_DETERMINACION_ID = " & DETERMINACION
'                Set rs2 = datos_bd(CO)
'                If rs2.RecordCount > 0 Then
'                    Do
'                        DETALLE = DETALLE & Format(rs2(0), "@@@@@@@@@@@@@@@@@@@@") & " : " & moneda(rs2(1)) & vbNewLine
'                        rs2.MoveNext
'                    Loop Until rs2.EOF
'                End If
'
'                DETALLE = DETALLE & vbNewLine
'                DETALLE = DETALLE & " Tarifa Nueva " & vbNewLine
'                DETALLE = DETALLE & " -------------- " & vbNewLine
'                If tarifas.ListItems.Count > 0 Then
'                  For i = 1 To tarifas.ListItems.Count
'                    If Trim(tarifas.ListItems(i).SubItems(2)) <> "" Then
'                      DETALLE = DETALLE & Format(tarifas.ListItems(i).SubItems(1), "@@@@@@@@@@@@@@@@@@@@") & " : " & moneda(tarifas.ListItems(i).SubItems(2)) & vbNewLine
'                    End If
'                  Next
'                End If
            
'                oParametro.Carga PARAM_USUARIO_VIGILADO_CORREO, ""
'                ret = Enviar_Mail_CDO(oParametro.getVALOR, asunto, DETALLE, vbNullString)
'            End If
'      End If
      insertar_configuracion DETERMINACION
      If USUARIO.getPER_FACTURACION = True Then
        Dim oTP As New clsTarifas_precios
        If PK <> 0 Then
          oTP.Eliminar_por_determinacion (PK)
        End If
        If tarifas.ListItems.Count > 0 Then
          For i = 1 To tarifas.ListItems.Count
              If Trim(tarifas.ListItems(i).SubItems(2)) <> "" Then
                  With oTP
                      .setTIPO_DETERMINACION_ID = DETERMINACION
                      .setTARIFA_ID = tarifas.ListItems(i).Text
                      .setPRECIO = moneda_bd(tarifas.ListItems(i).SubItems(2))
                      .Insertar
                  End With
              End If
          Next
        End If
      End If
      ' Equipos
      Dim OTDE As New clsTipos_determinacion_equipos
      OTDE.Eliminar DETERMINACION
      For i = 1 To listaEquipos.ListItems.Count
        With OTDE
            .setTIPO_DETERMINACION_ID = DETERMINACION
            .setEQUIPO_ID = listaEquipos.ListItems(i).Text
            .setORDEN = i
            .Insertar
        End With
      Next
      ' Reactivos
      Dim oTDB As New clsTipos_determinacion_botes_ex
      oTDB.Eliminar DETERMINACION
      For i = 1 To listaReactivos.ListItems.Count
        With oTDB
            .setTIPO_DETERMINACION_ID = DETERMINACION
            .setBOTE_EX_ID = listaReactivos.ListItems(i).Text
            .setTIPO = listaReactivos.ListItems(i).SubItems(3)
            .setORDEN = i
            .Insertar
        End With
      Next
        Me.MousePointer = 0
      
      If PK = 0 Then
          MsgBox "La determinación se ha introducido correctamente.", vbOKOnly + vbInformation, App.Title
      Else
          MsgBox "La determinación se ha modificado correctamente.", vbOKOnly + vbInformation, App.Title
      End If
      Unload Me
    End If

   On Error GoTo 0
   Exit Sub

cmdok_Click_Error:
        Me.MousePointer = 0

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure cmdok_Click of Formulario frmTD_Detalle"
End Sub
Private Sub Form_Load()
    log (Me.Name)
    cargar_botones Me
    Call permisos
    cabecera
    cargar_tarifas
    llenar_combo cmbFormula, New clsFormulas, 0, frmFORMULA_Detalle, ""
    llenar_combo cmbFamilia, New clsTipos_determinacion_familias, 0, Me, ""
    llenar_combo cmbtarifa, New clsTarifas_codigos, 0, Me, ""
    llenar_combo cmbPNT, New clsCa_documentos, 0, frmCA_Documento, " NOMBRE LIKE '%PNT%' "
    llenar_combo cmbnorma, New clsCa_normas, 0, frmCA_Normas, ""
    
    cargar_combos
    inicializar_grid
    If PK <> 0 Then
        lbltitulo = "Modificación del Tipo de determinación"
        cargar_td
        cargar_configuracion (PK)
        If chkSubcontratable.Value = Checked Then
            marcoSubcontratacion.Enabled = True
        End If
    Else
        lbltitulo = "Alta de nuevo Tipo de determinación"
        txtDatos(18) = "10"
    End If
    tarifa_modificada = False
    tabControl.Item(0).Selected = True
End Sub

Private Sub listaEquipos_DblClick()
    If listaEquipos.ListItems.Count > 0 Then
        frmEquipoEdicion.PK = listaEquipos.ListItems(listaEquipos.selectedItem.Index).Text
        frmEquipoEdicion.Show 1
    End If
End Sub


Private Sub listaSubcontratas_Click()
    If listaSubcontratas.ListItems.Count > 0 Then
       cmbSubcontratas.MostrarElemento listaSubcontratas.selectedItem.Text
       txtDatos(9).Text = listaSubcontratas.ListItems(listaSubcontratas.selectedItem.Index).SubItems(2)
       txtDatos(12).Text = listaSubcontratas.ListItems(listaSubcontratas.selectedItem.Index).SubItems(3)
       txtDatos(19).Text = listaSubcontratas.ListItems(listaSubcontratas.selectedItem.Index).SubItems(4)
    End If
End Sub


Private Sub tarifas_Click()
    If tarifas.ListItems.Count > 0 Then
        lblCampos(8) = "Precio Tarifa " & Trim(tarifas.ListItems(tarifas.selectedItem.Index).SubItems(1))
         txttarifa = Trim(tarifas.ListItems(tarifas.selectedItem.Index).SubItems(2))
         txttarifa.SetFocus
    End If
End Sub

Private Sub txtdatos_GotFocus(Index As Integer)
    txtDatos(Index).BackColor = &H80C0FF
    txtDatos(Index).SelStart = 0
    txtDatos(Index).SelLength = Len(txtDatos(Index))
End Sub

Private Sub txtdatos_KeyPress(Index As Integer, KeyAscii As Integer)
    If Index = 8 Or Index = 16 Then
        If KeyAscii = 46 Then
           KeyAscii = 44
        End If
    End If
End Sub

Private Sub txtdatos_LostFocus(Index As Integer)
    txtDatos(Index).BackColor = vbWhite
    If Index = 8 Then
        txtDatos(Index) = moneda(txtDatos(Index))
    End If
End Sub
Private Sub cargar_td()
    Dim oDET As New clsTipos_determinacion
    oDET.CargarTipoDeterminacion (PK)
    txtDatos(0) = oDET.getNOMBRE
    txtDatos(7) = oDET.getNOMBRE_INGLES
    txtDatos(1) = oDET.getDESCRIPCION
    txtDatos(2) = oDET.getPNT
    txtDatos(3) = oDET.getPROC_REF_EADS
    cmbPNT.MostrarElemento oDET.getPNT_VINCULADO
    cmbtarifa.MostrarElemento oDET.getTARIFA_CODIGO_ID
    chkrevisarfactura.Value = oDET.getREVISAR_FACTURA
    txtDatos(8) = moneda(oDET.getPRECIO)
    chkSubcontratable.Value = oDET.getES_SUBCONTRATABLE
    txtDatos(9).Text = ""
    txtDatos(12).Text = ""
    cargar_subcontratistas
    ' Formula
    cmbFormula.MostrarElemento oDET.getFORMULA_ID
    cmbFamilia.MostrarElemento oDET.getFAMILIA_ID
    txtDatos(15) = oDET.getMETODO
    cmbnorma.MostrarElemento oDET.getNORMA_ID
    ' Multitarifa
    Dim oMT As New clsTarifas_precios
    Dim rs As ADODB.Recordset
    Set rs = oMT.Listado_por_determinacion(PK)
    If rs.RecordCount <> 0 Then
        Dim i As Integer
        Do
                For i = 1 To tarifas.ListItems.Count
                    If CInt(tarifas.ListItems(i).Text) = CInt(rs("TARIFA_ID")) Then
                        tarifas.ListItems(i).SubItems(2) = moneda(CStr(rs("PRECIO")))
                    End If
                Next
            rs.MoveNext
        Loop Until rs.EOF
    End If
    chkparticulas.Value = oDET.getPARTICULAS
    ' Equipos
    Dim OTDEQUIPOS As New clsTipos_determinacion_equipos
    Set rs = OTDEQUIPOS.Listado(PK)
    If rs.RecordCount > 0 Then
        Do
            With listaEquipos.ListItems.Add(, , rs(0))
                .SubItems(1) = rs(1)
                .SubItems(2) = rs(2)
            End With
            rs.MoveNext
        Loop Until rs.EOF
    End If
    ' Reactivos
    Dim OTDR As New clsTipos_determinacion_botes_ex
    Dim oReactivo As New clsBotes_ex
    Dim oTb As New clsTipos_bote_ex
    Dim oTR As New clsTipos_reactivo_ex
    
    Dim oRPR As New clsRpr_botes
    Dim oTRPR As New clsRPR_Tipos
    Set rs = OTDR.Listado(PK)
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
                    .SubItems(1) = oTRPR.getNOMBRE
                    .SubItems(2) = Format(oRPR.getFECHA_CADUCIDAD, "DD-MM-YYYY")
                    .SubItems(3) = "I"
                End With
            End If
            rs.MoveNext
        Loop Until rs.EOF
    End If
End Sub
Private Function validar() As Boolean
    validar = True
    If Trim(txtDatos(0)) = "" Then
        MsgBox "Debe darle un nombre a la determinación.", vbInformation, App.Title
        validar = False
        Exit Function
    End If
    If cmbFormula.getPK_SALIDA = 0 Then
        MsgBox "Debe introducir una formula.", vbInformation, App.Title
        validar = False
        Exit Function
    End If
    ' Validar campos numéricos en el grid de configuración
    Dim i As Integer
    For i = gridTabla.LowerBound(1) To gridTabla.UpperBound(1)
        If Trim(gridTabla.Value(i, COLS.C_RANGO)) <> "" Then
            If Trim(gridTabla.Value(i, COLS.C_DIF_DUPLICADOS)) <> "" Then
                If Not IsNumeric(gridTabla.Value(i, COLS.C_DIF_DUPLICADOS)) Then
                    validar = False
                    MsgBox "El valor de % de Dif. entre duplicados debe ser numérico. Valor : " & gridTabla.Value(i, COLS.C_DIF_DUPLICADOS), vbExclamation, App.Title
                    Exit Function
                End If
            End If
            If Trim(gridTabla.Value(i, COLS.C_INCERTIDUMBRE)) <> "" Then
                If Not IsNumeric(gridTabla.Value(i, COLS.C_INCERTIDUMBRE)) Then
                    validar = False
                    MsgBox "El valor de la INCERTIDUMBRE debe ser numérico. Valor : " & gridTabla.Value(i, COLS.C_INCERTIDUMBRE), vbExclamation, App.Title
                    Exit Function
                End If
            End If
            If Trim(gridTabla.Value(i, COLS.C_DIF_AVISO)) <> "" Then
                If Not IsNumeric(gridTabla.Value(i, COLS.C_DIF_AVISO)) Then
                    validar = False
                    MsgBox "El valor de diferencia de % de Aviso de Rango debe ser numérico. Valor : " & gridTabla.Value(i, COLS.C_DIF_AVISO), vbExclamation, App.Title
                    Exit Function
                End If
            End If
        End If
    Next
End Function

Private Sub permisos()
    If USUARIO.getPER_FACTURACION = False Then
        Frame2.visible = False
    End If
End Sub
Private Sub cargar_tarifas()
    Dim oTarifa As New clsTarifas
    Dim rs As ADODB.Recordset
    Set rs = oTarifa.Listado_por_nombre
    If rs.RecordCount <> 0 Then
        Do
            With tarifas.ListItems.Add(, , rs(3))
                .SubItems(1) = rs(0)
                .SubItems(2) = " "
            End With
            rs.MoveNext
        Loop Until rs.EOF
    End If
End Sub

Private Sub txttarifa_GotFocus()
    txttarifa.SelStart = 0
    txttarifa.SelLength = Len(txttarifa.Text)
End Sub

Private Sub txttarifa_KeyPress(KeyAscii As Integer)
    If KeyAscii = 46 Then
       KeyAscii = 44
    End If
    If KeyAscii = 13 Then
        anadir_precio
    End If
End Sub

Private Sub cargar_combos()
'M0927-I
    llenar_combo cmbSubcontratas, New clsProveedor, 0, frmProveedores_Detalle, ""
'M0927-F
    llenar_combo cmbEquipos, New clsEquipos, 0, frmEquipoEdicion, ""
' NATALIA CORREO 25/05/2018
'    llenar_combo cmbReactivos, New clsBotes_ex, 0, Me, ""
'    llenar_combo cmbReactivosInternos, New clsRpr_botes, 0, Me, ""
    llenar_combo cmbReactivos, New clsBotes_ex, 0, Me, " AND FINALIZADO = 0 "
    llenar_combo cmbReactivosInternos, New clsRpr_botes, 0, Me, " AND ISNULL(fecha_fin)"
    
End Sub

'M0927-I
Private Sub cargar_subcontratistas()
    Dim oTDContratas As New clsTipos_determinacion_contratas
    Dim oProveedor As New clsProveedor
    Dim rsContratas As New ADODB.Recordset
    Set rsContratas = oTDContratas.Listado_Tipo(PK)
    
    If rsContratas.RecordCount > 0 Then
       Do
          With listaSubcontratas.ListItems.Add(, , rsContratas("CONTRATA_ID"))
            oProveedor.Carga rsContratas("CONTRATA_ID")
            .SubItems(1) = oProveedor.getNOMBRE
            .SubItems(2) = rsContratas("VALOR_REFERENCIA")
            .SubItems(3) = rsContratas("NORMATIVA_APLICABLE")
            .SubItems(4) = moneda(rsContratas("PRECIO"))
          End With
            
          rsContratas.MoveNext
       Loop Until rsContratas.EOF
    End If
    
    Set oTDContratas = Nothing
    Set oProveedor = Nothing
    Set rsContratas = Nothing
    
End Sub
'M0927-F
Private Sub inicializar_grid()
    gridTabla.ReDim 0, filas, 0, Col
    gridTabla.Clear
    Set grid.Array = gridTabla
    grid.Refresh
End Sub
Private Sub cargar_configuracion(TIPO_DETERMINACION_ID As Long)
    Dim oCONF As New clsTipos_determinacion_conf
    Dim rs As ADODB.Recordset
    Set rs = oCONF.Listado(TIPO_DETERMINACION_ID)
    If rs.RecordCount > 0 Then
        Dim i As Integer
        i = 0
        Do
            gridTabla(i, COLS.C_ORDEN) = texto(rs("ORDEN"))
            gridTabla(i, COLS.C_RANGO) = texto(rs("RANGO"))
            gridTabla(i, COLS.C_DIF_DUPLICADOS) = texto(rs("DIF_DUPLICADOS"))
            gridTabla(i, COLS.C_DIF_DUPLICADOS_NUMERICA) = texto(rs("DIF_DUPLICADOS_NUMERICA"))
            gridTabla(i, COLS.C_DIF_AVISO) = texto(rs("DIF_AVISO"))
            gridTabla(i, COLS.C_INCERTIDUMBRE) = texto(rs("INCERTIDUMBRE"))
            gridTabla(i, COLS.C_DIF_HISTORICO) = texto(rs("DIF_HISTORICO"))
            gridTabla(i, COLS.C_C_VARIACION) = texto(rs("C_VARIACION"))
            i = i + 1
            rs.MoveNext
        Loop Until rs.EOF
    End If
    Set rs = Nothing
End Sub
Private Sub insertar_configuracion(TIPO_DETERMINACION_ID As Long)
    Dim oCONF As New clsTipos_determinacion_conf
   On Error GoTo insertar_servicios_Error

    oCONF.Eliminar TIPO_DETERMINACION_ID
    Dim i As Integer
    For i = gridTabla.LowerBound(1) To gridTabla.UpperBound(1)
        If Trim(gridTabla.Value(i, COLS.C_RANGO)) <> "" Then
            With oCONF
                .setTIPO_DETERMINACION_ID = TIPO_DETERMINACION_ID
                .setORDEN = i
                .setRANGO = gridTabla.Value(i, COLS.C_RANGO)
                .setDIF_DUPLICADOS = gridTabla.Value(i, COLS.C_DIF_DUPLICADOS)
                .setDIF_DUPLICADOS_NUMERICA = gridTabla.Value(i, COLS.C_DIF_DUPLICADOS_NUMERICA)
                If gridTabla.Value(i, COLS.C_DIF_AVISO) = "" Then
                    .setDIF_AVISO = 10
                Else
                    .setDIF_AVISO = gridTabla.Value(i, COLS.C_DIF_AVISO)
                End If
                .setINCERTIDUMBRE = gridTabla.Value(i, COLS.C_INCERTIDUMBRE)
                .setDIF_HISTORICO = gridTabla.Value(i, COLS.C_DIF_HISTORICO)
                .setC_VARIACION = gridTabla.Value(i, COLS.C_C_VARIACION)
                .Insertar
            End With
        End If
    Next
    Set oCONF = Nothing

   On Error GoTo 0
   Exit Sub

insertar_servicios_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure insertar_configuracion of Formulario frmTD_Detalle"
End Sub


