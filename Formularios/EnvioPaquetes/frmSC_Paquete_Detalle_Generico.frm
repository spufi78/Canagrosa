VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{CEA3B7F6-2847-4E5E-A551-DB7A62489D44}#46.0#0"; "miCombo.ocx"
Object = "{77DDF307-D82B-4757-8B3A-106EC9D75D4B}#8.0#0"; "tdbg8.ocx"
Begin VB.Form frmSC_Paquete_Detalle_Generico 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   9480
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   14220
   Icon            =   "frmSC_Paquete_Detalle_Generico.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9480
   ScaleWidth      =   14220
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdFacturacion 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Facturas"
      Height          =   915
      Left            =   9045
      Style           =   1  'Graphical
      TabIndex        =   46
      ToolTipText     =   "Enviar informe por E-mail"
      Top             =   8550
      Width           =   1275
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
      ForeColor       =   &H80000002&
      Height          =   285
      Index           =   2
      Left            =   11430
      Locked          =   -1  'True
      MaxLength       =   100
      TabIndex        =   41
      Top             =   5670
      Width           =   2730
   End
   Begin VB.Frame frmHistoria 
      BackColor       =   &H00C0C0C0&
      Height          =   4110
      Left            =   4320
      TabIndex        =   16
      Top             =   2070
      Width           =   6315
      Begin VB.CommandButton Command1 
         Caption         =   "Cerrar"
         Height          =   465
         Left            =   2475
         TabIndex        =   21
         Top             =   3510
         Width           =   1455
      End
      Begin VB.Frame Frame5 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Recepción"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1005
         Left            =   45
         TabIndex        =   20
         Top             =   2430
         Width           =   6225
         Begin VB.TextBox txtFechaRecepcion 
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
            ForeColor       =   &H00000000&
            Height          =   285
            Left            =   1125
            Locked          =   -1  'True
            MaxLength       =   100
            TabIndex        =   33
            TabStop         =   0   'False
            Top             =   585
            Width           =   1380
         End
         Begin VB.TextBox txtUsuarioRecepcion 
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
            Height          =   285
            Left            =   1125
            Locked          =   -1  'True
            TabIndex        =   28
            Top             =   225
            Width           =   4920
         End
         Begin VB.Label Label8 
            BackColor       =   &H00C0C0C0&
            Caption         =   "Fecha:"
            Height          =   195
            Left            =   495
            TabIndex        =   27
            Top             =   630
            Width           =   510
         End
         Begin VB.Label Label5 
            BackColor       =   &H00C0C0C0&
            Caption         =   "Usuario:"
            Height          =   195
            Left            =   405
            TabIndex        =   24
            Top             =   270
            Width           =   645
         End
      End
      Begin VB.Frame Frame4 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Trámite"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000001&
         Height          =   1005
         Left            =   45
         TabIndex        =   19
         Top             =   1440
         Width           =   6225
         Begin VB.TextBox txtFechaTramite 
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
            ForeColor       =   &H00000000&
            Height          =   285
            Left            =   1125
            Locked          =   -1  'True
            MaxLength       =   100
            TabIndex        =   32
            TabStop         =   0   'False
            Top             =   540
            Width           =   1380
         End
         Begin VB.TextBox txtUsuarioTramite 
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
            Height          =   285
            Left            =   1125
            Locked          =   -1  'True
            TabIndex        =   29
            Top             =   180
            Width           =   4920
         End
         Begin VB.Label Label7 
            BackColor       =   &H00C0C0C0&
            Caption         =   "Fecha:"
            Height          =   195
            Left            =   495
            TabIndex        =   26
            Top             =   630
            Width           =   510
         End
         Begin VB.Label Label4 
            BackColor       =   &H00C0C0C0&
            Caption         =   "Usuario:"
            Height          =   195
            Left            =   405
            TabIndex        =   23
            Top             =   270
            Width           =   645
         End
      End
      Begin VB.Frame Frame3 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Petición"
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
         Height          =   1005
         Left            =   45
         TabIndex        =   18
         Top             =   450
         Width           =   6225
         Begin VB.TextBox txtFechaPeticion 
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
            ForeColor       =   &H00000000&
            Height          =   285
            Left            =   1125
            Locked          =   -1  'True
            MaxLength       =   100
            TabIndex        =   31
            TabStop         =   0   'False
            Top             =   585
            Width           =   1380
         End
         Begin VB.TextBox txtUsuarioPeticion 
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
            Height          =   285
            Left            =   1125
            Locked          =   -1  'True
            TabIndex        =   30
            Top             =   225
            Width           =   4920
         End
         Begin VB.Label Label6 
            BackColor       =   &H00C0C0C0&
            Caption         =   "Fecha:"
            Height          =   195
            Left            =   495
            TabIndex        =   25
            Top             =   630
            Width           =   510
         End
         Begin VB.Label Label3 
            BackColor       =   &H00C0C0C0&
            Caption         =   "Usuario:"
            Height          =   195
            Left            =   405
            TabIndex        =   22
            Top             =   270
            Width           =   645
         End
      End
      Begin VB.Label lblSubtitulo 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFC0C0&
         Caption         =   "Historia"
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
         Height          =   285
         Index           =   0
         Left            =   45
         TabIndex        =   17
         Top             =   135
         Width           =   6225
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00C0C0C0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   0
      TabIndex        =   35
      Top             =   6030
      Width           =   14145
      Begin VB.CommandButton cmdAceptar 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Height          =   465
         Left            =   13500
         Style           =   1  'Graphical
         TabIndex        =   36
         ToolTipText     =   "Ver muestra seleccionada"
         Top             =   180
         Width           =   510
      End
      Begin pryCombo.miCombo cmbConceptos 
         Height          =   330
         Left            =   1440
         TabIndex        =   37
         Top             =   270
         Width           =   11760
         _ExtentX        =   20743
         _ExtentY        =   582
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Añadir Concepto"
         Height          =   195
         Index           =   3
         Left            =   135
         TabIndex        =   38
         Top             =   315
         Width           =   1185
      End
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00C0E0FF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   285
      Left            =   0
      MaxLength       =   100
      TabIndex        =   15
      TabStop         =   0   'False
      Text            =   "Importe Total Presupuestado: "
      Top             =   5670
      Width           =   11400
   End
   Begin VB.TextBox txtDatos 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H00000000&
      Height          =   1365
      Index           =   3
      Left            =   45
      MaxLength       =   255
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   13
      Top             =   7155
      Width           =   14115
   End
   Begin VB.CommandButton cmdAdjuntos 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Adjuntos"
      Height          =   915
      Left            =   10350
      Style           =   1  'Graphical
      TabIndex        =   11
      ToolTipText     =   "Enviar informe por E-mail"
      Top             =   8550
      Width           =   1275
   End
   Begin VB.CommandButton cmdEliminar 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Eliminar Concepto"
      Height          =   870
      Left            =   45
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   8595
      Width           =   1545
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Datos subcontratación"
      Height          =   1725
      Left            =   0
      TabIndex        =   4
      Top             =   315
      Width           =   14145
      Begin pryCombo.miCombo cmbClientes 
         Height          =   375
         Left            =   1035
         TabIndex        =   55
         Top             =   1305
         Width           =   5430
         _ExtentX        =   9578
         _ExtentY        =   661
      End
      Begin VB.TextBox txtEdicion 
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
         ForeColor       =   &H00C00000&
         Height          =   330
         Left            =   8910
         MaxLength       =   100
         TabIndex        =   49
         TabStop         =   0   'False
         Top             =   630
         Width           =   750
      End
      Begin VB.CheckBox chkTramite 
         Caption         =   "Check1"
         Height          =   240
         Left            =   8100
         TabIndex        =   44
         Top             =   990
         Visible         =   0   'False
         Width           =   240
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
         ForeColor       =   &H00000000&
         Height          =   285
         Index           =   1
         Left            =   1035
         MaxLength       =   100
         TabIndex        =   0
         TabStop         =   0   'False
         Top             =   270
         Width           =   1650
      End
      Begin MSComCtl2.DTPicker datFecha 
         Height          =   315
         Left            =   3825
         TabIndex        =   1
         TabStop         =   0   'False
         Top             =   225
         Width           =   1590
         _ExtentX        =   2805
         _ExtentY        =   556
         _Version        =   393216
         Enabled         =   0   'False
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
         Format          =   60424193
         CurrentDate     =   2
         MinDate         =   2
      End
      Begin pryCombo.miCombo cmbUsuario 
         Height          =   330
         Left            =   8910
         TabIndex        =   2
         TabStop         =   0   'False
         Top             =   270
         Width           =   5070
         _ExtentX        =   8943
         _ExtentY        =   582
      End
      Begin pryCombo.miCombo cmbSubcontratas 
         Height          =   330
         Left            =   1035
         TabIndex        =   40
         Top             =   585
         Width           =   5415
         _ExtentX        =   9551
         _ExtentY        =   582
      End
      Begin pryCombo.miCombo cmbSubtipo 
         Height          =   375
         Left            =   1035
         TabIndex        =   48
         Top             =   945
         Width           =   5430
         _ExtentX        =   9578
         _ExtentY        =   661
      End
      Begin MSDataListLib.DataCombo cmbCentro 
         Bindings        =   "frmSC_Paquete_Detalle_Generico.frx":08CA
         Height          =   315
         Left            =   11070
         TabIndex        =   51
         Top             =   630
         Width           =   1875
         _ExtentX        =   3307
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
      Begin MSDataListLib.DataCombo cmbMoneda 
         Height          =   315
         Left            =   11070
         TabIndex        =   53
         Top             =   990
         Width           =   1875
         _ExtentX        =   3307
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         Style           =   2
         ListField       =   "0"
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
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Cliente"
         Height          =   195
         Index           =   8
         Left            =   450
         TabIndex        =   56
         Top             =   1350
         Width           =   480
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Moneda"
         Height          =   195
         Index           =   7
         Left            =   10395
         TabIndex        =   54
         Top             =   1035
         Width           =   585
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Centro"
         Height          =   195
         Index           =   22
         Left            =   10395
         TabIndex        =   52
         Top             =   675
         Width           =   465
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Edición"
         Height          =   195
         Index           =   5
         Left            =   8100
         TabIndex        =   50
         Top             =   675
         Width           =   525
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Subtipo"
         Height          =   195
         Index           =   4
         Left            =   405
         TabIndex        =   47
         Top             =   990
         Width           =   540
      End
      Begin VB.Label lblNecesita 
         BackColor       =   &H00C0C0C0&
         Caption         =   "No necesita trámite"
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
         Left            =   8415
         TabIndex        =   45
         Top             =   990
         Visible         =   0   'False
         Width           =   1770
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Usuario"
         Height          =   195
         Index           =   2
         Left            =   8100
         TabIndex        =   8
         Top             =   315
         Width           =   540
      End
      Begin VB.Label Label2 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Fecha"
         Height          =   240
         Left            =   3195
         TabIndex        =   7
         Top             =   270
         Width           =   510
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Subcontrata"
         Height          =   195
         Index           =   1
         Left            =   90
         TabIndex        =   6
         Top             =   630
         Width           =   870
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Código SC"
         Height          =   195
         Index           =   0
         Left            =   90
         TabIndex        =   5
         Top             =   315
         Width           =   750
      End
   End
   Begin VB.CommandButton cmdok 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Modificar"
      Height          =   915
      Left            =   11655
      Style           =   1  'Graphical
      TabIndex        =   12
      ToolTipText     =   "Modificar paquete"
      Top             =   8550
      Width           =   1275
   End
   Begin VB.CommandButton cmdcancel 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Cancelar"
      Height          =   915
      Left            =   12960
      Style           =   1  'Graphical
      TabIndex        =   14
      ToolTipText     =   "Salir"
      Top             =   8550
      Width           =   1230
   End
   Begin TrueDBGrid80.TDBGrid grid 
      Height          =   3630
      Left            =   0
      TabIndex        =   39
      Top             =   2070
      Width           =   14145
      _ExtentX        =   24950
      _ExtentY        =   6403
      _LayoutType     =   4
      _RowHeight      =   -2147483647
      _WasPersistedAsPixels=   0
      Columns(0)._VlistStyle=   0
      Columns(0)._MaxComboItems=   5
      Columns(0).Caption=   "REF."
      Columns(0).DataField=   ""
      Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(1)._VlistStyle=   0
      Columns(1)._MaxComboItems=   5
      Columns(1).Caption=   "DESCRIPCIÓN"
      Columns(1).DataField=   ""
      Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(2)._VlistStyle=   0
      Columns(2)._MaxComboItems=   5
      Columns(2).Caption=   "UDs."
      Columns(2).DataField=   ""
      Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(3)._VlistStyle=   0
      Columns(3)._MaxComboItems=   5
      Columns(3).Caption=   "DESC. (%)"
      Columns(3).DataField=   ""
      Columns(3).NumberFormat=   "Standard"
      Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(4)._VlistStyle=   0
      Columns(4)._MaxComboItems=   5
      Columns(4).Caption=   "PVP (Ud.)"
      Columns(4).DataField=   ""
      Columns(4).NumberFormat=   "Standard"
      Columns(4).ConvertEmptyCell=   1
      Columns(4)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(5)._VlistStyle=   0
      Columns(5)._MaxComboItems=   5
      Columns(5).Caption=   "IMPORTE"
      Columns(5).DataField=   ""
      Columns(5).NumberFormat=   "Standard"
      Columns(5)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns.Count   =   6
      Splits(0)._UserFlags=   0
      Splits(0).PartialRightColumn=   0   'False
      Splits(0).MarqueeStyle=   1
      Splits(0).AllowRowSizing=   0   'False
      Splits(0).RecordSelectors=   0   'False
      Splits(0).RecordSelectorWidth=   503
      Splits(0).AllowColSelect=   0   'False
      Splits(0).AlternatingRowStyle=   -1  'True
      Splits(0).DividerColor=   12632256
      Splits(0).SpringMode=   0   'False
      Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
      Splits(0)._ColumnProps(0)=   "Columns.Count=6"
      Splits(0)._ColumnProps(1)=   "Column(0).Width=2302"
      Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
      Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=2223"
      Splits(0)._ColumnProps(4)=   "Column(0)._EditAlways=0"
      Splits(0)._ColumnProps(5)=   "Column(0)._ColStyle=131585"
      Splits(0)._ColumnProps(6)=   "Column(0).Order=1"
      Splits(0)._ColumnProps(7)=   "Column(1).Width=11986"
      Splits(0)._ColumnProps(8)=   "Column(1).DividerColor=0"
      Splits(0)._ColumnProps(9)=   "Column(1)._WidthInPix=11906"
      Splits(0)._ColumnProps(10)=   "Column(1)._EditAlways=0"
      Splits(0)._ColumnProps(11)=   "Column(1)._ColStyle=131585"
      Splits(0)._ColumnProps(12)=   "Column(1).Order=2"
      Splits(0)._ColumnProps(13)=   "Column(2).Width=1429"
      Splits(0)._ColumnProps(14)=   "Column(2).DividerColor=0"
      Splits(0)._ColumnProps(15)=   "Column(2)._WidthInPix=1349"
      Splits(0)._ColumnProps(16)=   "Column(2)._EditAlways=0"
      Splits(0)._ColumnProps(17)=   "Column(2)._ColStyle=131585"
      Splits(0)._ColumnProps(18)=   "Column(2).Order=3"
      Splits(0)._ColumnProps(19)=   "Column(3).Width=1720"
      Splits(0)._ColumnProps(20)=   "Column(3).DividerColor=0"
      Splits(0)._ColumnProps(21)=   "Column(3)._WidthInPix=1640"
      Splits(0)._ColumnProps(22)=   "Column(3)._EditAlways=0"
      Splits(0)._ColumnProps(23)=   "Column(3)._ColStyle=131585"
      Splits(0)._ColumnProps(24)=   "Column(3).Order=4"
      Splits(0)._ColumnProps(25)=   "Column(4).Width=2646"
      Splits(0)._ColumnProps(26)=   "Column(4).DividerColor=0"
      Splits(0)._ColumnProps(27)=   "Column(4)._WidthInPix=2566"
      Splits(0)._ColumnProps(28)=   "Column(4)._EditAlways=0"
      Splits(0)._ColumnProps(29)=   "Column(4)._ColStyle=131585"
      Splits(0)._ColumnProps(30)=   "Column(4).Order=5"
      Splits(0)._ColumnProps(31)=   "Column(5).Width=4022"
      Splits(0)._ColumnProps(32)=   "Column(5).DividerColor=0"
      Splits(0)._ColumnProps(33)=   "Column(5)._WidthInPix=3942"
      Splits(0)._ColumnProps(34)=   "Column(5)._EditAlways=0"
      Splits(0)._ColumnProps(35)=   "Column(5)._ColStyle=131585"
      Splits(0)._ColumnProps(36)=   "Column(5).Order=6"
      Splits.Count    =   1
      PrintInfos(0)._StateFlags=   3
      PrintInfos(0).Name=   "piInternal 0"
      PrintInfos(0).PageHeaderFont=   "Size=8.25,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=MS Sans Serif"
      PrintInfos(0).PageFooterFont=   "Size=8.25,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=MS Sans Serif"
      PrintInfos(0).PageHeaderHeight=   0
      PrintInfos(0).PageFooterHeight=   0
      PrintInfos.Count=   1
      AllowDelete     =   -1  'True
      AllowAddNew     =   -1  'True
      ColumnFooters   =   -1  'True
      DataMode        =   4
      DefColWidth     =   0
      HeadLines       =   1
      FootLines       =   0
      RowDividerStyle =   0
      TabAction       =   2
      WrapCellPointer =   -1  'True
      MultipleLines   =   0
      EmptyRows       =   -1  'True
      CellTips        =   2
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
      _StyleDefs(6)   =   "Style:id=1,.parent=0,.namedParent=37,.bgcolor=&HDEEDFA&,.fgcolor=&H0&,.bold=0"
      _StyleDefs(7)   =   ":id=1,.fontsize=825,.italic=0,.underline=0,.strikethrough=0,.charset=0"
      _StyleDefs(8)   =   ":id=1,.fontname=MS Sans Serif"
      _StyleDefs(9)   =   "CaptionStyle:id=4,.parent=2,.namedParent=41,.alignment=0,.fgcolor=&H80000001&"
      _StyleDefs(10)  =   ":id=4,.bold=-1,.fontsize=975,.italic=0,.underline=0,.strikethrough=0,.charset=0"
      _StyleDefs(11)  =   ":id=4,.fontname=MS Sans Serif"
      _StyleDefs(12)  =   "HeadingStyle:id=2,.parent=1,.namedParent=38,.bgcolor=&H8080FF&,.fgcolor=&H0&"
      _StyleDefs(13)  =   ":id=2,.bold=0,.fontsize=975,.italic=0,.underline=0,.strikethrough=0,.charset=0"
      _StyleDefs(14)  =   ":id=2,.fontname=MS Sans Serif"
      _StyleDefs(15)  =   "FooterStyle:id=3,.parent=1,.namedParent=39,.bgcolor=&H80000009&"
      _StyleDefs(16)  =   ":id=3,.fgcolor=&H80000001&,.bold=0,.fontsize=975,.italic=0,.underline=0"
      _StyleDefs(17)  =   ":id=3,.strikethrough=0,.charset=0"
      _StyleDefs(18)  =   ":id=3,.fontname=MS Sans Serif"
      _StyleDefs(19)  =   "InactiveStyle:id=5,.parent=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(20)  =   "SelectedStyle:id=6,.parent=1,.namedParent=40,.bgcolor=&HFFFFFF&"
      _StyleDefs(21)  =   "EditorStyle:id=7,.parent=1,.bgcolor=&HFFFFFF&,.fgcolor=&HFFFFFF&"
      _StyleDefs(22)  =   "HighlightRowStyle:id=8,.parent=1,.namedParent=42"
      _StyleDefs(23)  =   "EvenRowStyle:id=9,.parent=1,.namedParent=43,.alignment=3"
      _StyleDefs(24)  =   "OddRowStyle:id=10,.parent=1,.namedParent=44"
      _StyleDefs(25)  =   "RecordSelectorStyle:id=45,.parent=2,.namedParent=47"
      _StyleDefs(26)  =   "FilterBarStyle:id=48,.parent=1,.namedParent=50"
      _StyleDefs(27)  =   "Splits(0).Style:id=11,.parent=1"
      _StyleDefs(28)  =   "Splits(0).CaptionStyle:id=20,.parent=4"
      _StyleDefs(29)  =   "Splits(0).HeadingStyle:id=12,.parent=2,.namedParent=38"
      _StyleDefs(30)  =   "Splits(0).FooterStyle:id=13,.parent=3"
      _StyleDefs(31)  =   "Splits(0).InactiveStyle:id=14,.parent=5"
      _StyleDefs(32)  =   "Splits(0).SelectedStyle:id=16,.parent=6,.namedParent=40"
      _StyleDefs(33)  =   "Splits(0).EditorStyle:id=15,.parent=7,.namedParent=40"
      _StyleDefs(34)  =   "Splits(0).HighlightRowStyle:id=17,.parent=8"
      _StyleDefs(35)  =   "Splits(0).EvenRowStyle:id=18,.parent=9,.namedParent=43,.alignment=2"
      _StyleDefs(36)  =   "Splits(0).OddRowStyle:id=19,.parent=10,.namedParent=44"
      _StyleDefs(37)  =   "Splits(0).RecordSelectorStyle:id=46,.parent=45"
      _StyleDefs(38)  =   "Splits(0).FilterBarStyle:id=49,.parent=48"
      _StyleDefs(39)  =   "Splits(0).Columns(0).Style:id=24,.parent=11"
      _StyleDefs(40)  =   "Splits(0).Columns(0).HeadingStyle:id=21,.parent=12"
      _StyleDefs(41)  =   "Splits(0).Columns(0).FooterStyle:id=22,.parent=13"
      _StyleDefs(42)  =   "Splits(0).Columns(0).EditorStyle:id=23,.parent=15"
      _StyleDefs(43)  =   "Splits(0).Columns(1).Style:id=28,.parent=11"
      _StyleDefs(44)  =   "Splits(0).Columns(1).HeadingStyle:id=25,.parent=12"
      _StyleDefs(45)  =   "Splits(0).Columns(1).FooterStyle:id=26,.parent=13"
      _StyleDefs(46)  =   "Splits(0).Columns(1).EditorStyle:id=27,.parent=15"
      _StyleDefs(47)  =   "Splits(0).Columns(2).Style:id=32,.parent=11"
      _StyleDefs(48)  =   "Splits(0).Columns(2).HeadingStyle:id=29,.parent=12"
      _StyleDefs(49)  =   "Splits(0).Columns(2).FooterStyle:id=30,.parent=13"
      _StyleDefs(50)  =   "Splits(0).Columns(2).EditorStyle:id=31,.parent=15"
      _StyleDefs(51)  =   "Splits(0).Columns(3).Style:id=54,.parent=11"
      _StyleDefs(52)  =   "Splits(0).Columns(3).HeadingStyle:id=51,.parent=12"
      _StyleDefs(53)  =   "Splits(0).Columns(3).FooterStyle:id=52,.parent=13"
      _StyleDefs(54)  =   "Splits(0).Columns(3).EditorStyle:id=53,.parent=15"
      _StyleDefs(55)  =   "Splits(0).Columns(4).Style:id=36,.parent=11"
      _StyleDefs(56)  =   "Splits(0).Columns(4).HeadingStyle:id=33,.parent=12"
      _StyleDefs(57)  =   "Splits(0).Columns(4).FooterStyle:id=34,.parent=13"
      _StyleDefs(58)  =   "Splits(0).Columns(4).EditorStyle:id=35,.parent=15"
      _StyleDefs(59)  =   "Splits(0).Columns(5).Style:id=58,.parent=11"
      _StyleDefs(60)  =   "Splits(0).Columns(5).HeadingStyle:id=55,.parent=12"
      _StyleDefs(61)  =   "Splits(0).Columns(5).FooterStyle:id=56,.parent=13"
      _StyleDefs(62)  =   "Splits(0).Columns(5).EditorStyle:id=57,.parent=15"
      _StyleDefs(63)  =   "Named:id=37:Normal"
      _StyleDefs(64)  =   ":id=37,.parent=0,.alignment=2,.bgcolor=&H80000018&,.bold=0,.fontsize=825"
      _StyleDefs(65)  =   ":id=37,.italic=0,.underline=0,.strikethrough=0,.charset=0"
      _StyleDefs(66)  =   ":id=37,.fontname=MS Sans Serif"
      _StyleDefs(67)  =   "Named:id=38:Heading"
      _StyleDefs(68)  =   ":id=38,.parent=37,.valignment=2,.bgcolor=&HC0C0C0&,.fgcolor=&H80000012&"
      _StyleDefs(69)  =   ":id=38,.wraptext=-1,.bold=-1,.fontsize=825,.italic=0,.underline=0"
      _StyleDefs(70)  =   ":id=38,.strikethrough=0,.charset=0"
      _StyleDefs(71)  =   ":id=38,.fontname=MS Sans Serif"
      _StyleDefs(72)  =   "Named:id=39:Footing"
      _StyleDefs(73)  =   ":id=39,.parent=37,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(74)  =   "Named:id=40:Selected"
      _StyleDefs(75)  =   ":id=40,.parent=37,.bgcolor=&H8080FF&,.fgcolor=&H0&,.bold=0,.fontsize=825"
      _StyleDefs(76)  =   ":id=40,.italic=0,.underline=0,.strikethrough=0,.charset=0"
      _StyleDefs(77)  =   ":id=40,.fontname=MS Sans Serif"
      _StyleDefs(78)  =   "Named:id=41:Caption"
      _StyleDefs(79)  =   ":id=41,.parent=38,.alignment=2"
      _StyleDefs(80)  =   "Named:id=42:HighlightRow"
      _StyleDefs(81)  =   ":id=42,.parent=37,.bgcolor=&H80000008&,.fgcolor=&H80000005&"
      _StyleDefs(82)  =   "Named:id=43:EvenRow"
      _StyleDefs(83)  =   ":id=43,.parent=37,.bgcolor=&HFFFFFF&"
      _StyleDefs(84)  =   "Named:id=44:OddRow"
      _StyleDefs(85)  =   ":id=44,.parent=37,.bgcolor=&HE9E9E9&"
      _StyleDefs(86)  =   "Named:id=47:RecordSelector"
      _StyleDefs(87)  =   ":id=47,.parent=38"
      _StyleDefs(88)  =   "Named:id=50:FilterBar"
      _StyleDefs(89)  =   ":id=50,.parent=37"
   End
   Begin VB.Label lblTramite 
      BackColor       =   &H80000009&
      Caption         =   "Necesita permisos como tramitador si desea modificarlo"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   240
      Left            =   1935
      TabIndex        =   43
      Top             =   8820
      Visible         =   0   'False
      Width           =   4785
   End
   Begin VB.Label lblCampos 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      Caption         =   "Total Importe"
      Height          =   195
      Index           =   6
      Left            =   9810
      TabIndex        =   42
      Top             =   5715
      Width           =   930
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "F5 - HISTORIA"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   240
      Left            =   12555
      TabIndex        =   34
      Top             =   45
      Width           =   1530
   End
   Begin VB.Label lblObservaciones 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Observaciones:"
      Height          =   285
      Left            =   45
      TabIndex        =   10
      Top             =   6840
      Width           =   1140
   End
   Begin VB.Label lblSubtitulo 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Caption         =   "Detalle de la subcontratación"
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
      Height          =   285
      Index           =   2
      Left            =   0
      TabIndex        =   3
      Top             =   0
      Width           =   14190
   End
End
Attribute VB_Name = "frmSC_Paquete_Detalle_Generico"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'---------------------------------------------------------------------------------------------'
' MANTIS 1163: Nueva ventana
'---------------------------------------------------------------------------------------------'

Option Explicit
Public PK As Long
'M1274-i
Public EDICION As Long
'M1274-F

Private x As New XArrayDB
Private fila As Integer

Const filas As Integer = 100
Const Col As Integer = 5
Const cReferencia As Integer = 0
Const cDescripcion As Integer = 1
Const cUnidades As Integer = 2
Const cDescuento As Integer = 3
Const cPrecio As Integer = 4
Const cImporte As Integer = 5

Private Sub cmbMoneda_Change()
    SumarImportes
End Sub
'M1257-I
Private Sub cmdFacturacion_Click()
        If cmbSubcontratas.getTEXTO <> "" Then
            With frmProveedores_Facturas
                .TOBJETO = TOBJETO.TOBJETO_SC_GENERICA
                .COBJETO = PK
                .PK = cmbSubcontratas.getPK_SALIDA
                .Show 1
            End With
        End If
End Sub

'M1257-F
Private Sub Form_Load()
    'Ventana y grid
    inicializar_ventana
    cargar_botones Me
    
    'Combos
    Call cargar_combo_subcontratas
    chkTramite.Value = 1
    chkTramite.Enabled = False
    llenar_combo cmbUsuario, New clsUsuarios, 0, frmUsuarios, ""
    llenar_combo cmbConceptos, New clsSc_paquetes_detalle_generico, 0, Me, ""
    cargar_combo cmbCentro, New clsCentros
    llenar_combo cmbClientes, New clsCliente, 0, frmClientes, ""

    cmbUsuario.desactivar
    txtDatos(2).Enabled = False
'M1257-I
    Dim oDeco As New clsDecodificadora
    oDeco.cargar_mi_combo cmbSubtipo, DECODIFICADORA.SC_SUBTIPOS
    oDeco.cargar_combo cmbMoneda, DECODIFICADORA.DECODIFICADORA_MONEDA
    Set oDeco = Nothing

'   datFechaFactura.value = Date
'M1257-F
    frmHistoria.visible = False
    txtedicion.Enabled = False
    If PK <> 0 Then
        'Carga el formulario completo
        'M1171-I
        If Not USUARIO.getPER_TRAMITACION_CONTRATA Then
            cmdok.Enabled = False
            cmdEliminar.Enabled = False
            lblTramite.visible = True
        End If
        'M1171-F
        MODIFICACION
    Else
        'Habilita los controles y carga valores por defecto
        'M1171-I
        lblTramite.visible = False
        lblNecesita.visible = True
        chkTramite.visible = True
        cmdAdjuntos.Enabled = False
        txtedicion.Text = 1
        'M1171-F
        Alta
    End If
    If Not USUARIO.getPER_TESORERIA_FP Then
        cmdFacturacion.visible = False
    End If
    
End Sub

Public Sub inicializar_ventana()
    Dim i As Integer
    log (Me.Name)
    Me.top = 1700
    Me.Left = 300
    fila = 0
    x.ReDim 0, filas, 0, Col
    x.Clear
    Set grid.Array = x
End Sub

Private Sub cmdAdjuntos_Click()
    With frmAdjuntos
        .TOBJETO = TOBJETO.TOBJETO_PAQUETE_SUBCONTRATA
        .COBJETO = PK
        .Show 1
    End With
    Set frmAdjuntos = Nothing
End Sub

Private Sub cmdEliminar_Click()

    If MsgBox("¿Desea eliminar el concepto?", vbQuestion + vbYesNo, App.Title) = vbYes Then
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
        SumarImportes
    End If
End Sub

Private Sub Command1_Click()
    frmHistoria.visible = False
End Sub


Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    On Error GoTo Form_KeyUp_Error

    Me.MousePointer = 0
    Select Case KeyCode
        Case 116 ' F5 Datos especiales
            If frmHistoria.visible = False Then
                Frame3.ForeColor = SC_COLOR_PENDIENTE
                Frame4.ForeColor = SC_COLOR_TRAMITADO
                Frame5.ForeColor = SC_COLOR_RECIBIDO
                frmHistoria.visible = True
            Else
                frmHistoria.visible = False
            End If
    End Select

   On Error GoTo 0
   Exit Sub
Form_KeyUp_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure Form_KeyUp of Formulario frmEmpleados_Matriz"

End Sub

' botones
'ALTA Y MODIFICACION
Private Sub cmdok_Click()
    If PK <> 0 Then
        Dim oSC As New clsSC_Paquetes
        oSC.Carga PK, EDICION
        If oSC.getTIPO = TOBJETO_SC_PEACH Then
            MsgBox "No se puede modificar. Es una SC con origen PEACH.", vbCritical, App.Title
            Exit Sub
        End If
    End If
    guardarCambios
End Sub


Private Sub guardarCambios()

   On Error GoTo modificarPaquete_Error
    Dim strMensaje As String
    Dim FECHAHORA As Date
    If datos_correctos Then
    
        FECHAHORA = Now
        Dim oSC_Paquete As New clsSC_Paquetes
        Dim lngPaquete As Long
        With oSC_Paquete
'M1274-I
               If PK <> 0 Then
                  .Carga PK, EDICION
                  If MsgBox("¿Desea generar una nueva edición de la subcontratacion", vbYesNo + vbQuestion, App.Title) = vbYes Then
                     EDICION = EDICION + 1
                     PK = 0
                  End If
               Else
                  EDICION = 1
               End If
               .setEDICION = EDICION
'M1274-F
              .setCENTRO_ID = cmbCentro.BoundText
              .setPRESUPUESTO = Replace(Format(Replace(txtDatos(2), cmbMoneda.Text, ""), "0.00"), ",", ".")
              .setOBSERVACIONES = txtDatos(3)
              .setSUBCONTRATA_ID = cmbSubcontratas.getPK_SALIDA
              .setCLIENTE_ID = cmbClientes.getPK_SALIDA
              .setSUBTIPO = cmbSubtipo.getPK_SALIDA

              .setFACTURA_RECIBIDA = 0
              .setFFACTURA = "0000-00-00"
              .setNFACTURA = 0
              .setMONEDA = cmbMoneda.BoundText
              
              If PK = 0 Then
                 strMensaje = "Se va a crear un nuevo paquete. ¿Está seguro?"
                  
                 .setFECHA_CREACION = Left(Format(FECHAHORA, "yyyy-mm-dd hh:nn:ss"), 10)
                 .setHORA_CREACION = Right(Format(FECHAHORA, "yyyy-mm-dd hh:nn:ss"), 8)
                 .setUSUARIO_ID = USUARIO.getID_EMPLEADO
                 'M1171-I
                 '.setAPROBADOR_ID = 0
                    '  .setESTADO = SC_ESTADO_PENDIENTE
                    If chkTramite.Value = 0 Then
                       .setESTADO = SC_ESTADO_PENDIENTE
                                .setAPROBADOR_ID = 0
                       .setFECHA_APROBACION = "0000-00-00"
                    Else
                       .setESTADO = SC_ESTADO_TRAMITADO
                       .setAPROBADOR_ID = USUARIO.getID_EMPLEADO
                       .setFECHA_APROBACION = Format(Date, "yyyy-mm-dd")
                    End If
                 'M1171-F
                 .setTIPO = TOBJETO_SC_GENERICA
                 'M1257-I
                 .setSUBTIPO = CLng(cmbSubtipo.getPK_SALIDA)
                 'M1257-F
                 'M1274-i
                 If Trim(.getFECHA_RECEPCION) = "" Then
                    .setFECHA_RECEPCION = "0000-00-00"
                 End If
                 'M1274-F
              Else
                 strMensaje = "Va a modificar el registro. ¿Está seguro?"
              End If
              
              If MsgBox(strMensaje, vbQuestion + vbYesNo, App.Title) = vbYes Then
                If PK <> 0 Then
                   lngPaquete = PK
                   
                   If .Modificar(lngPaquete, EDICION) = False Then
                    MsgBox "Se ha producido un error al modificar el registro.", vbCritical, App.Title
                    Exit Sub
                   End If
                Else
                   lngPaquete = .Insertar
                   If lngPaquete = 0 Then
                    MsgBox "Se ha producido un error al crear la nueva subcontratación.", vbCritical, App.Title
                    Exit Sub
                   End If
                End If
                        ' CONCEPTOS
                Dim opd As New clsSc_paquetes_detalle_generico
                Dim i As Long
                opd.Eliminar_Paquete lngPaquete, oSC_Paquete.getEDICION
                For i = 0 To filas - 1
                   If filaCargada(i) Then
                       opd.setPAQUETE_ID = lngPaquete
                       opd.setREFERENCIA = Trim(x(i, cReferencia))
                       opd.setDESCRIPCION = Trim(x(i, cDescripcion))
                       opd.setUNIDADES = CLng(x(i, cUnidades))
                       opd.setDESCUENTO = CLng(x(i, cDescuento))
                       opd.setPRECIO = moneda_bd(Trim(x(i, cPrecio)))
                       opd.setIMPORTE = moneda_bd(Trim(x(i, cImporte)))
                       opd.setEDICION = oSC_Paquete.getEDICION
                       opd.Insertar
                   End If
                Next i
                'JGM
                If frmSC_Listado.lstPaquetes.ListItems.Count > 0 Then
                    frmSC_Listado.lstPaquetes.selectedItem.SubItems(frmSC_Listado.COL_EDICION) = EDICION
                End If
                If PK = 0 Then
                   'ENVÍO DE CORREO
                   If chkTramite.Value = 0 And Not USUARIO.getPER_TRAMITACION_CONTRATA Then
                       envioCorreoTramite lngPaquete, EDICION
                   End If
                   MsgBox "El código de subcontratación nº " & oSC_Paquete.getCODIGO_SC & " Edición: " & EDICION & " se ha creado correctamente.", vbOKOnly + vbInformation, App.Title
                Else
                   MsgBox "El registro se ha modificado correctamente.", vbOKOnly + vbInformation, App.Title
                End If
             
              End If
        End With
        Unload Me
      End If
     
   On Error GoTo 0
   Exit Sub

modificarPaquete_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure modificarPaquete of Formulario frmSC_Paquete_Detalle_Generico"
End Sub
 
Private Function recorrer_filas() As Double
    
    Dim indice As Integer
    Dim encontrado As Boolean
    encontrado = True
    indice = 0
    recorrer_filas = 0
    Do
        If x(indice, cPrecio) <> "" Then
           recorrer_filas = recorrer_filas + CDbl(Trim(x(indice, cPrecio)))
        Else
           encontrado = False
        End If
        indice = indice + 1
    Loop Until Not encontrado Or indice > filas
End Function

Private Sub cmdcancel_Click()
    Unload Me
End Sub
' --------------------------

Private Sub cmdAceptar_Click()
    On Error GoTo fallo
    Dim oConcepto As New clsSc_paquetes_detalle_generico
    Dim i As Integer
    Dim pos As Integer
    
    ' Cargamos los datos del concepto
    If oConcepto.Carga(CLng(cmbConceptos.getPK_SALIDA)) = True Then
        pos = calcularNumeroFilas()
        x(pos, cReferencia) = CStr(oConcepto.getREFERENCIA)
        x(pos, cDescripcion) = CStr(oConcepto.getDESCRIPCION)
        x(pos, cPrecio) = moneda(Trim(oConcepto.getPRECIO))
        x(pos, cImporte) = moneda(Trim(oConcepto.getPRECIO))
        x(pos, cUnidades) = 1
        x(pos, cDescuento) = 0
        
        grid.Row = 0
        grid.Col = 0
        grid.Refresh
        grid.SetFocus
    Else
        MsgBox "Error al cargar el documento.", vbInformation, App.Title
    End If
    Set oConcepto = Nothing
    
    SumarImportes
    Exit Sub
fallo:
    MsgBox "Error al cargar los datos del documento. " & Err.Description, vbCritical, App.Title
End Sub

Private Function calcularNumeroFilas() As Integer
    Dim i As Long
    i = 0
    Do While filaCargada(i)
        i = i + 1
    Loop
    calcularNumeroFilas = i  'sale con i incrementada sobre la última fila real
End Function

Private Function filaCargada(fila As Long) As Boolean
    Dim i As Integer
    filaCargada = False
    For i = 0 To Col - 1
        If Trim(x(fila, i)) <> "" Then
            filaCargada = True
        End If
    Next i
End Function

Private Sub lblNecesita_Click()
    If chkTramite.Value = 0 Then
       chkTramite.Value = 1
    Else
       chkTramite.Value = 0
    End If
End Sub

Private Sub txtdatos_KeyPress(Index As Integer, KeyAscii As Integer)
    Select Case Index
        Case 2, 3:
            
        Case Else
            KeyAscii = 0
    End Select
End Sub

' funciones auxiliares del formulario
Public Sub Alta()
    
    Me.MousePointer = vbHourglass
    cmdok.Caption = "Alta"
    txtDatos(1) = "N/A"
    lblsubtitulo(2) = "Creación de nueva subcontratación"
    cmbSubcontratas.activar
    cmbUsuario.activar
    cmbUsuario.MostrarElemento USUARIO.getID_EMPLEADO
'M1257-I
'    txtFactura = ""
'M1257-F
    txtDatos(2) = ""
    Me.MousePointer = vbNormal
    
End Sub

Public Sub MODIFICACION()
    Dim oSC_Paquete As New clsSC_Paquetes
    Dim usu As New clsUsuarios
    Dim rs As ADODB.Recordset
    Dim lngTotalConceptosPaquete As Long
    
    Me.MousePointer = vbHourglass
    cmbSubcontratas.activar
'M1274-I
'    If oSC_Paquete.Carga(PK) = True Then
    If oSC_Paquete.Carga(PK, EDICION) = True Then
'M1274-F
        With oSC_Paquete
        'M1274-I
            txtedicion.Text = EDICION
        'M1274-F
            cmbCentro.BoundText = .getCENTRO_ID
            cmbMoneda.BoundText = .getMONEDA
            txtDatos(1) = .getCODIGO_SC
            lblsubtitulo(2) = "Detalle del paquete: " & .getCODIGO_SC
            txtDatos(2) = .getPRESUPUESTO
            txtDatos(3) = .getOBSERVACIONES
            cmbSubcontratas.MostrarElemento .getSUBCONTRATA_ID
            cmbClientes.MostrarElemento .getCLIENTE_ID
            cmbUsuario.MostrarElemento .getUSUARIO_ID
            cmbSubtipo.MostrarElemento .getSUBTIPO
            'M1257-I
            'txtFactura = .getNFACTURA
            'If .getFFACTURA <> "" Then
            '    datFechaFactura = .getFFACTURA
            'End If
            'M1257-F
            
            'Carga del histórico
            If IsDate(.getFECHA_CREACION) Then
                datFecha = .getFECHA_CREACION
                txtFechaPeticion.Text = Format(.getFECHA_CREACION, "yyyy-mm-dd")
                If usu.CARGAR(.getUSUARIO_ID) Then
                  txtUsuarioPeticion.Text = usu.getNOMBRE & " " & usu.getAPELLIDOS
                Else
                   txtUsuarioPeticion.Text = "N/A"
                End If
                Else
                txtFechaPeticion.Text = " -- "
                txtUsuarioPeticion.Text = "#Error en la fecha de recepción#"
                End If
            If IsDate(.getFECHA_APROBACION) Then
                txtFechaTramite.Text = Format(.getFECHA_APROBACION, "yyyy-mm-dd")
                If usu.CARGAR(.getAPROBADOR_ID) Then
                    txtUsuarioTramite.Text = usu.getNOMBRE & " " & usu.getAPELLIDOS
                Else
                    txtUsuarioTramite.Text = "N/A"
                End If
            Else
                txtFechaTramite.Text = "N/A"
                txtUsuarioTramite.Text = "N/A"
            End If
            If IsDate(.getFECHA_RECEPCION) Then
                txtFechaRecepcion.Text = Format(.getFECHA_RECEPCION, "yyyy-mm-dd")
                If usu.CARGAR(.getRECEPTOR_ID) Then
                    txtUsuarioRecepcion.Text = usu.getNOMBRE & " " & usu.getAPELLIDOS
                Else
                    txtUsuarioRecepcion.Text = "N/A"
                End If
            Else
                txtFechaRecepcion.Text = "N/A"
                txtUsuarioRecepcion.Text = "N/A"
            End If
            
            'Carga del GRID con listado de conceptos
            Dim oConceptos As New clsSc_paquetes_detalle_generico
            Set rs = oConceptos.Listado_conceptos(.getID_PAQUETE, EDICION)
            lngTotalConceptosPaquete = rs.RecordCount
            
            If rs.RecordCount <> 0 Then
                Dim i As Integer
                Dim impTxt As String
                i = 0
                Do
                    x(i, cReferencia) = CStr(rs(2))
                    x(i, cDescripcion) = CStr(rs(3))
                    x(i, cUnidades) = CStr(rs(4))
                    x(i, cDescuento) = CStr(rs(5))
                    x(i, cPrecio) = moneda(Trim(rs(6)))
                    If rs(7) = 0 Then
                        impTxt = moneda(Trim(calcularImporte(CInt(rs(4)), CInt(rs(5)), CDbl(rs(6)))))
                    Else
                        impTxt = moneda(Trim(rs(7)))
                    End If
                    x(i, cImporte) = impTxt
                    i = i + 1
                    grid.Row = 0
                    grid.Col = 0
                    grid.Refresh
                   ' grid.SetFocus
                    rs.MoveNext
                Loop Until rs.EOF
            End If
        End With
    End If
    Me.MousePointer = vbNormal
    SC_bloqueaModificacion (oSC_Paquete.getESTADO)
    Set oSC_Paquete = Nothing
    SumarImportes
    lblsubtitulo(2) = lblsubtitulo(2) & " - Nº CONCEPTOS: " & lngTotalConceptosPaquete
    
End Sub
Public Sub SC_bloqueaModificacion(estadoPaquete As Long)
    If Not USUARIO.getPER_TRAMITACION_CONTRATA Then
        Select Case estadoPaquete
        Case SC_ESTADO_RECIBIDO
            cmdEliminar.Enabled = False
            cmdok.Enabled = False
            Frame2.Enabled = False
        Case SC_ESTADO_HISTORICO
            cmdEliminar.Enabled = False
            cmdok.Enabled = False
            Frame2.Enabled = False
        End Select
    End If
End Sub
Private Function calcularImporte(unidades As Integer, desc As Integer, PRECIO As Double) As String
    Dim importeTotal As Double
    Dim DESCUENTO As Double
    
    If unidades = 0 Then
        calcularImporte = moneda(0)
    End If
    
    DESCUENTO = 0
    If desc > 0 Then
       If desc > 100 Then
          desc = 100
       End If
       DESCUENTO = (unidades * PRECIO * desc) / 100
    End If
    
    importeTotal = (unidades * PRECIO) - DESCUENTO
    calcularImporte = moneda(CStr(importeTotal))
End Function

Private Sub envioCorreoTramite(ID As Long, EDICION As Long)
'-------------- ENVÍO AUTOMÁTICO DE CORREO A LISTA DE DISTRIBUCIÓN TRAS TRAMITACIÓN DE PAQUETE ------'
    Dim destinatario As String
    Dim mensaje As String
    Dim ASUNTO As String
    Dim oParametro As New clsParametros
    Dim oPaquete As New clsSC_Paquetes
    Dim i As Integer, numFilas As Integer
    
    oParametro.Carga parametros.PARAM_CORREO_DISTRIBUCION_TRAMITE, ""
    oPaquete.Carga ID, EDICION
    destinatario = oParametro.getVALOR
    
    If destinatario <> "" Then
        
        ASUNTO = "Tramitación de pedido a proveedor. Código : " & oPaquete.getCODIGO_SC
        mensaje = "Se ha creado el siguiente pedido a proveedor: " & vbNewLine & vbNewLine

        mensaje = mensaje & vbNewLine & " Código : " & oPaquete.getCODIGO_SC
        mensaje = mensaje & vbNewLine & " Presupuesto : " & oPaquete.getPRESUPUESTO
        mensaje = mensaje & vbNewLine & " Generada por : " & "(" & USUARIO.getUSUARIO & ") " & USUARIO.getNOMBRE & " " & USUARIO.getAPELLIDOS
        
        ' LISTADO DE CONCEPTOS
        i = 0
        numFilas = calcularNumeroFilas
        If numFilas > 0 Then
             mensaje = mensaje & vbNewLine
             mensaje = mensaje & vbNewLine & "--------------------------------------------------------------------------------------------------------"
             mensaje = mensaje & vbNewLine & "                  ** LISTA DE CONCEPTOS **"
             mensaje = mensaje & vbNewLine & "--------------------------------------------------------------------------------------------------------"
             Do
                 mensaje = mensaje & vbNewLine
                 mensaje = mensaje & Format(Left(x(i, cDescripcion), 50), "!" & String(50, "@")) & "    "
                 mensaje = mensaje & Format(Left(x(i, cUnidades), 4), "!" & String(4, "@")) & "    "
                 mensaje = mensaje & Format(Left(x(i, cDescuento), 4), "!" & String(4, "@")) & "    "
                 mensaje = mensaje & Format(Left(x(i, cPrecio), 8), "!" & String(8, "@")) & "    "
                 mensaje = mensaje & Format(Left(x(i, cImporte), 8), "!" & String(8, "@")) & "    "
                 i = i + 1
             Loop Until i > numFilas
        End If

        mensaje = mensaje & vbNewLine
        mensaje = mensaje & vbNewLine
        mensaje = mensaje & " Mensaje enviado automáticamente desde Geslab el " & Format(Date, "dd-mm-yyyy") & " a las " & Format(Time, "hh:mm:ss")
        
       ret = Enviar_Mail_CDO(destinatario, ASUNTO, mensaje, vbNullString)
      '  ret = Enviar_Mail_CDO("daniel.gallardo@ixitec.net", ASUNTO, mensaje, vbNullString)
    End If
    Set oParametro = Nothing
End Sub

Private Function datos_correctos() As Boolean
    datos_correctos = True

    If cmbSubcontratas.getPK_SALIDA = 0 Then
        MsgBox "Debe indicar la subcontrata antes de generar el paquete", vbExclamation, App.Title
        datos_correctos = False
        cmbSubcontratas.SetFocus
        Exit Function
    End If
    If cmbUsuario.getTEXTO = "" Then
        MsgBox "Debe indicar el usuario antes de generar el paquete", vbExclamation, App.Title
        datos_correctos = False
        cmbUsuario.SetFocus
        Exit Function
    End If

    If Trim(txtDatos(2)) = "" Then ' presupuesto
        If MsgBox("No ha indicado ningún presupuesto. ¿Modificar el paquete sin presupuesto?", vbYesNo + vbInformation, App.Title) = vbNo Then
        datos_correctos = False
        txtDatos(2).SetFocus
        Exit Function
        End If
    End If
    If cmbCentro.BoundText = "" Then
        MsgBox "No ha indicado el centro.", vbExclamation, App.Title
        datos_correctos = False
        cmbCentro.SetFocus
        Exit Function
    End If
    If cmbMoneda.BoundText = "" Then
        MsgBox "No ha indicado la moneda.", vbExclamation, App.Title
        datos_correctos = False
        cmbMoneda.SetFocus
        Exit Function
    End If

End Function

Private Sub cargar_combo_subcontratas()
    llenar_combo cmbSubcontratas, New clsProveedor, 0, frmProveedores_Detalle, " ES_SUBCONTRATA = 1 "
End Sub

Private Sub SumarImportes()

   Dim indice As Integer
   Dim Suma As Double
   Suma = 0

   For indice = 0 To filas - 1
        If IsNumeric(x(indice, cImporte)) Then
            Suma = Suma + CDbl(x(indice, cImporte))
        End If
   Next indice
   txtDatos(2) = Replace(moneda(CStr(Suma)), "€", cmbMoneda.Text)
End Sub

Private Sub grid_RowColChange(LastRow As Variant, ByVal LastCol As Integer)

    Select Case LastCol
    Case cUnidades To cImporte
        If Not IsNumeric(x(LastRow, cPrecio)) Then
            x(LastRow, cPrecio) = "0"
        End If

        If Not IsNumeric(x(LastRow, cUnidades)) Then
            x(LastRow, cUnidades) = "0"
        End If
        
        If Not IsNumeric(x(LastRow, cDescuento)) Then
            x(LastRow, cDescuento) = "0"
        End If
        
        x(LastRow, cImporte) = calcularImporte(CInt(x(LastRow, cUnidades)), CInt(x(LastRow, cDescuento)), CDbl(x(LastRow, cPrecio)))
        SumarImportes
    End Select
    grid.Refresh

    
End Sub
