VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{CEA3B7F6-2847-4E5E-A551-DB7A62489D44}#39.0#0"; "miCombo.ocx"
Begin VB.Form frmOferta_Nueva 
   BackColor       =   &H00C0C0C0&
   Caption         =   "Gestión de Ofertas"
   ClientHeight    =   9060
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   13560
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmOferta_Nueva.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   9060
   ScaleWidth      =   13560
   StartUpPosition =   2  'CenterScreen
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
      Left            =   5310
      Picture         =   "frmOferta_Nueva.frx":08CA
      Style           =   1  'Graphical
      TabIndex        =   40
      Top             =   8100
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
      Left            =   4005
      Picture         =   "frmOferta_Nueva.frx":1194
      Style           =   1  'Graphical
      TabIndex        =   34
      Top             =   8100
      Width           =   1275
   End
   Begin VB.TextBox txtdatos 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0FF&
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
      Left            =   8415
      TabIndex        =   32
      Top             =   8460
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
      Left            =   2700
      Picture         =   "frmOferta_Nueva.frx":1A5E
      Style           =   1  'Graphical
      TabIndex        =   20
      Top             =   8100
      Width           =   1275
   End
   Begin VB.CommandButton cmdanadir 
      BackColor       =   &H00E0E0E0&
      Caption         =   "AñadirLínea"
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
      Picture         =   "frmOferta_Nueva.frx":2328
      Style           =   1  'Graphical
      TabIndex        =   18
      Top             =   8100
      Width           =   1275
   End
   Begin VB.CommandButton cmdMod 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Modificar Línea"
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
      Picture         =   "frmOferta_Nueva.frx":2BF2
      Style           =   1  'Graphical
      TabIndex        =   19
      Top             =   8100
      Width           =   1275
   End
   Begin VB.TextBox datos 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      DataField       =   "USUARIO"
      DataSource      =   "Adodc1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   1335
      Index           =   4
      Left            =   11385
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   17
      Top             =   6705
      Width           =   1695
   End
   Begin VB.TextBox datos 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      DataField       =   "USUARIO"
      DataSource      =   "Adodc1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   1335
      Index           =   3
      Left            =   9585
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   16
      Top             =   6705
      Width           =   1785
   End
   Begin VB.TextBox datos 
      Appearance      =   0  'Flat
      DataField       =   "USUARIO"
      DataSource      =   "Adodc1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   1350
      Index           =   2
      Left            =   4860
      MaxLength       =   255
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   15
      Top             =   6705
      Width           =   4700
   End
   Begin VB.TextBox datos 
      Appearance      =   0  'Flat
      DataField       =   "USUARIO"
      DataSource      =   "Adodc1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   1350
      Index           =   1
      Left            =   45
      MaxLength       =   255
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   14
      Top             =   6705
      Width           =   4800
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
      Left            =   11385
      Style           =   1  'Graphical
      TabIndex        =   21
      Top             =   8100
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
      Left            =   12465
      Style           =   1  'Graphical
      TabIndex        =   22
      Top             =   8100
      Width           =   1050
   End
   Begin MSComctlLib.ListView lista 
      Height          =   3720
      Left            =   45
      TabIndex        =   13
      Top             =   2925
      Width           =   13035
      _ExtentX        =   22992
      _ExtentY        =   6562
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      AllowReorder    =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   12640511
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
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   "Datos comúnes"
      ForeColor       =   &H80000008&
      Height          =   2535
      Left            =   45
      TabIndex        =   24
      Top             =   360
      Width           =   13470
      Begin VB.OptionButton opIdioma 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Oferta en Inglés"
         Height          =   195
         Index           =   1
         Left            =   2205
         TabIndex        =   48
         Top             =   2250
         Width           =   2130
      End
      Begin VB.OptionButton opIdioma 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Oferta en Español"
         Height          =   195
         Index           =   0
         Left            =   135
         TabIndex        =   47
         Top             =   2250
         Value           =   -1  'True
         Width           =   2130
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
         TabIndex        =   43
         Top             =   225
         Width           =   930
      End
      Begin VB.Frame frameSubtipo 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         Caption         =   "SubTipo Oferta"
         Enabled         =   0   'False
         ForeColor       =   &H80000008&
         Height          =   1545
         Left            =   9495
         TabIndex        =   35
         Top             =   675
         Visible         =   0   'False
         Width           =   2040
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
            TabIndex        =   39
            Top             =   315
            Width           =   1635
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
            TabIndex        =   38
            Top             =   630
            Width           =   1905
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
            TabIndex        =   37
            Top             =   945
            Width           =   1815
         End
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
            TabIndex        =   36
            Top             =   1260
            Width           =   1635
         End
      End
      Begin VB.TextBox txtdatos 
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
         TabIndex        =   12
         Text            =   "frmOferta_Nueva.frx":34BC
         Top             =   1530
         Width           =   5955
      End
      Begin VB.TextBox txtdatos 
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
         Height          =   465
         Index           =   0
         Left            =   1305
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   11
         Top             =   1035
         Width           =   5955
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
         Left            =   5400
         TabIndex        =   1
         Top             =   270
         Value           =   1  'Checked
         Width           =   1815
      End
      Begin VB.Frame Frame3 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         Caption         =   "Logos"
         ForeColor       =   &H80000008&
         Height          =   1545
         Left            =   11565
         TabIndex        =   30
         Top             =   675
         Width           =   1635
         Begin VB.CheckBox chkLogo 
            BackColor       =   &H00E0E0E0&
            Caption         =   "ENAC (Materiales)"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   3
            Left            =   135
            TabIndex        =   45
            Top             =   630
            Width           =   1455
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
            TabIndex        =   9
            Top             =   1260
            Width           =   780
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
            TabIndex        =   8
            Top             =   1035
            Width           =   1005
         End
         Begin VB.CheckBox chkLogo 
            BackColor       =   &H00E0E0E0&
            Caption         =   "ENAC (Agrícola)"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   0
            Left            =   135
            TabIndex        =   7
            Top             =   225
            Width           =   1455
         End
      End
      Begin VB.Frame Frame2 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         Caption         =   "Tipo Oferta"
         ForeColor       =   &H80000008&
         Height          =   1545
         Left            =   7650
         TabIndex        =   2
         Top             =   675
         Width           =   1815
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
            TabIndex        =   46
            Top             =   1215
            Width           =   1635
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
            TabIndex        =   6
            Top             =   990
            Width           =   1635
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
            TabIndex        =   5
            Top             =   765
            Width           =   1500
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
            TabIndex        =   4
            Top             =   540
            Width           =   1365
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
            TabIndex        =   3
            Top             =   315
            Value           =   -1  'True
            Width           =   1635
         End
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
         TabIndex        =   25
         Top             =   270
         Width           =   1290
      End
      Begin MSComCtl2.DTPicker fecha 
         Height          =   330
         Left            =   3645
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
         Format          =   60948481
         CurrentDate     =   38000
         MinDate         =   2
      End
      Begin pryCombo.miCombo cmbclientes 
         Height          =   345
         Left            =   1305
         TabIndex        =   10
         Top             =   675
         Width           =   5955
         _ExtentX        =   10504
         _ExtentY        =   609
      End
      Begin MSDataListLib.DataCombo cmbDatos 
         Height          =   360
         Index           =   1
         Left            =   8280
         TabIndex        =   41
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
         TabIndex        =   44
         Top             =   315
         Width           =   525
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
         TabIndex        =   29
         Top             =   1755
         Width           =   1065
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
         TabIndex        =   31
         Top             =   1170
         Width           =   390
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
         TabIndex        =   28
         Top             =   720
         Width           =   480
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
         TabIndex        =   27
         Top             =   315
         Width           =   555
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Fecha"
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
         TabIndex        =   26
         Top             =   315
         Width           =   450
      End
   End
   Begin VB.Image flecha 
      Height          =   480
      Index           =   1
      Left            =   13050
      Picture         =   "frmOferta_Nueva.frx":3511
      Top             =   4815
      Width           =   480
   End
   Begin VB.Image flecha 
      Height          =   480
      Index           =   0
      Left            =   13050
      Picture         =   "frmOferta_Nueva.frx":3A51
      Top             =   4005
      Width           =   480
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
      Left            =   8415
      TabIndex        =   33
      Top             =   8190
      Width           =   1245
   End
   Begin VB.Label lbltitulo 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      Caption         =   "Creación de Oferta"
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
      TabIndex        =   23
      Top             =   0
      Width           =   13455
   End
End
Attribute VB_Name = "frmOferta_Nueva"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim tipo As Integer
Public PK As Long
Public PK_EDICION As Integer
Public Nueva_Edicion As Boolean
Const campos = 4
Private Sub cmdAdjuntos_Click()
    With frmMuestras_Adjuntos
        .inicializar
        .tipo = ADJUNTOS_TIPOS.ADJUNTOS_TIPOS_OFERTAS
        .PK_OFERTA = PK
        .Show 1
    End With
End Sub

Private Sub cmdAnadir_Click()
    With lista.ListItems.Add(, , Replace(datos(1), vbNewLine, " "))
        .SubItems(1) = Replace(datos(2), vbNewLine, " ")
        .SubItems(2) = datos(3)
        .SubItems(3) = datos(4)
    End With
    lista.ListItems(lista.ListItems.Count).EnsureVisible
    borrar_campos
End Sub

Private Sub cmdBorrarLinea_Click()
    If lista.ListItems.Count > 0 Then
        lista.ListItems.Remove lista.selectedItem.Index
    End If
    borrar_campos
End Sub

Private Sub cmdcancel_Click()
    Unload Me
End Sub

Private Sub cmdCriterio_Click()
    frmOferta_Seleccion.TIPO_OFERTA = tipo
    frmOferta_Seleccion.Show 1
    calcular_total
End Sub

Private Sub cmdMod_Click()
    If lista.ListItems.Count > 0 Then
        lista.ListItems(lista.selectedItem.Index).Text = Replace(datos(1), vbNewLine, " ")
        lista.ListItems(lista.selectedItem.Index).SubItems(1) = Replace(datos(2), vbNewLine, " ")
'        If tipo = 0 Or tipo = 1 Then
            lista.ListItems(lista.selectedItem.Index).SubItems(2) = datos(3)
            lista.ListItems(lista.selectedItem.Index).SubItems(3) = datos(4)
'        Else
'            lista.ListItems(lista.SelectedItem.Index).SubItems(2) = datos(4)
'        End If
    End If
    borrar_campos
End Sub

Private Sub cmdok_Click()
   On Error GoTo cmdok_Click_Error
    Dim nueva As Boolean
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
            .setFECHA = Format(fecha.value, "dd-mm-yyyy")
            .setNUMERO = datos(0)
            If opIdioma(0).value = True Then
                .setIDIOMA = 0
            Else
                .setIDIOMA = 1
            End If
            If chkLogo(0).value = Checked Then
                .setLOGO_ENAC = 1
            End If
            If chkLogo(1).value = Checked Then
                .setLOGO_NADCAP = 1
            End If
            If chkLogo(2).value = Checked Then
                .setLOGO_EQUA = 1
            End If
            If chkLogo(3).value = Checked Then
                .setLOGO_ENACM = 1
            End If
            If chkSello.value = Checked Then
                .setSELLO = 1
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
                If opsubTipo(1).value = True Then
                    .setSUBTIPO_OFERTA = 1
                End If
                If opsubTipo(2).value = True Then
                    .setSUBTIPO_OFERTA = 2
                End If
                If opsubTipo(3).value = True Then
                    .setSUBTIPO_OFERTA = 3
                End If
                If opsubTipo(4).value = True Then
                    .setSUBTIPO_OFERTA = 4
                End If
            End If
            .setTOTAL = txtDatos(2)
            .setUSUARIO_ID = usuario.getID_EMPLEADO
            .setESTADO_OFERTA = cmbDatos(1).BoundText
        End With
        Dim OFERTA As Long
        If nueva Then
            OFERTA = oOferta.Insertar
        Else
            oOferta.Modificar PK, PK_EDICION
            oOferta_Detalle.Eliminar (PK)
            OFERTA = PK
        End If
        ' Detalle
        oOferta.Quitar_Ultima oOferta.getNUMERO
        Dim i As Integer
        Dim bano_anterior As String
        For i = 1 To lista.ListItems.Count
            With oOferta_Detalle
                .setOFERTA_ID = OFERTA
                .setEDICION = oOferta.getEDICION
'                If PK = 0 Then
'                    .setEDICION = 1
'                Else
'                    .setEDICION = datos(5) + 1
'                End If
                
                If Trim(lista.ListItems(i).Text) = "" Then
                    .setBANO = bano_anterior
                Else
                    .setBANO = Replace(lista.ListItems(i).Text, vbNewLine, " ")
                    bano_anterior = Replace(lista.ListItems(i).Text, vbNewLine, " ")
                End If
                .setDETERMINACION = Replace(lista.ListItems(i).SubItems(1), vbNewLine, " ")
'                If TIPO = 0 Or TIPO = 1 Then
                    .setRANGO = Replace(lista.ListItems(i).SubItems(2), vbNewLine, " ")
                    .setPRECIO = lista.ListItems(i).SubItems(3)
'                Else
'                    .setPRECIO = lista.ListItems(i).SubItems(2)
'                End If
                .setORDEN = i
                If Trim(.getBANO) = "" And Trim(.getDETERMINACION) = "" And Trim(.getRANGO) = "" And Trim(.getPRECIO) = "" Then
                    
                Else
                    .Insertar
                End If
            End With
        Next
        Me.MousePointer = 0
        MsgBox "La oferta ha sido almacenada correctamente.", vbInformation, App.Title
        Unload Me
    End If

   On Error GoTo 0
   Exit Sub

cmdok_Click_Error:
    Me.MousePointer = 0
    error_grave ("Error " & Err.Number & " (" & Err.Description & ") in procedure cmdok_Click of Formulario frmOferta_Nueva")
End Sub

Private Sub datos_KeyPress(Index As Integer, KeyAscii As Integer)
    If Index = 4 Then
       If KeyAscii = 46 Then
           KeyAscii = 44
        End If
    End If
End Sub

Private Sub datos_LostFocus(Index As Integer)
    If Index = 4 Then
        datos(4) = Format(datos(4), "currency")
    End If
End Sub

Private Sub flecha_Click(Index As Integer)
    Dim aux As String
    Dim i As Integer
    If lista.ListItems.Count > 0 Then
        If Index = 0 Then 'Subir
           If lista.selectedItem.Index > 1 Then
              aux = lista.ListItems(lista.selectedItem.Index - 1).Text
              lista.ListItems(lista.selectedItem.Index - 1).Text = lista.ListItems(lista.selectedItem.Index).Text
              lista.ListItems(lista.selectedItem.Index).Text = aux
              For i = 1 To campos - 1
                  aux = lista.ListItems(lista.selectedItem.Index - 1).SubItems(i)
                  lista.ListItems(lista.selectedItem.Index - 1).SubItems(i) = lista.ListItems(lista.selectedItem.Index).SubItems(i)
                  lista.ListItems(lista.selectedItem.Index).SubItems(i) = aux
              Next
              Set lista.selectedItem = lista.ListItems(lista.selectedItem.Index - 1)
           End If
        Else ' Bajar
           If lista.selectedItem.Index < lista.ListItems.Count Then
              aux = lista.ListItems(lista.selectedItem.Index + 1).Text
              lista.ListItems(lista.selectedItem.Index + 1).Text = lista.ListItems(lista.selectedItem.Index).Text
              lista.ListItems(lista.selectedItem.Index).Text = aux
              For i = 1 To campos - 1
                  aux = lista.ListItems(lista.selectedItem.Index + 1).SubItems(i)
                  lista.ListItems(lista.selectedItem.Index + 1).SubItems(i) = lista.ListItems(lista.selectedItem.Index).SubItems(i)
                  lista.ListItems(lista.selectedItem.Index).SubItems(i) = aux
              Next
              Set lista.selectedItem = lista.ListItems(lista.selectedItem.Index + 1)
           End If
        End If
    End If

End Sub

Private Sub Form_Load()
    log Me.Name
    cargar_botones Me
    cargar_combo
    opTipo_Click (0)
    fecha = Date
    If PK = 0 Then
        Dim oOferta As New clsOfertas
        datos(0) = oOferta.Calcular_Numero
        cmbDatos(1).BoundText = 0
        cmdAdjuntos.Enabled = False
        datos(5) = 1
    Else
        Frame2.Enabled = False
        lbltitulo = "Modificación de Oferta"
        lbltitulo.BackColor = &H80FF&
        cargar_oferta
        cmdAdjuntos.Enabled = True
    End If
    If usuario.getID_EMPLEADO = 7 Then
        chkSello.Enabled = False
    End If
End Sub

Public Sub cargar_combo()
    'Clientes
    llenar_combo cmbClientes, New clsCliente, 0, frmClientes, ""
    Dim oDec As New clsDecodificadora
    oDec.cargar_combo cmbDatos(1), decodificadora.ESTADOS_OFERTAS
End Sub

Private Sub lista_Click()
    If lista.ListItems.Count > 0 Then
        datos(1) = lista.ListItems(lista.selectedItem.Index).Text
        datos(2) = lista.ListItems(lista.selectedItem.Index).SubItems(1)
'        If tipo = 0 Or tipo = 1 Then
            datos(3) = lista.ListItems(lista.selectedItem.Index).SubItems(2)
            datos(4) = lista.ListItems(lista.selectedItem.Index).SubItems(3)
'        Else
'            datos(4) = lista.ListItems(lista.SelectedItem.Index).SubItems(2)
'        End If
        datos(1).SetFocus
    End If
End Sub

Public Sub cargar_oferta()
    If PK > 0 Then
        Dim oOferta As New clsOfertas
        With oOferta
            .Carga PK, PK_EDICION
            datos(0) = .getNUMERO
            datos(5) = .getEDICION
            cmbClientes.MostrarElemento .getCLIENTE_ID
            fecha = .getFECHA
            chkSello.value = .getSELLO
            chkLogo(0).value = .getLOGO_ENAC
            chkLogo(1).value = .getLOGO_NADCAP
            chkLogo(2).value = .getLOGO_EQUA
            chkLogo(3).value = .getLOGO_ENACM
            opIdioma(.getIDIOMA).value = True
            opTipo(.getTIPO_OFERTA).value = True
            If .getSUBTIPO_OFERTA <> 0 Then
                opsubTipo(.getSUBTIPO_OFERTA).value = True
            End If
            tipo = .getTIPO_OFERTA
            txtDatos(0) = .getPLAZO_ENTREGA
            txtDatos(1) = .getOBSERVACIONES
            txtDatos(2) = .getTOTAL
            cmbDatos(1).BoundText = .getESTADO_OFERTA
        End With
        ' Detalle
        Dim oOferta_Detalle As New clsOfertas_detalle
        Dim rs As ADODB.RecordSet
        Set rs = oOferta_Detalle.Listado(PK, PK_EDICION)
        If rs.RecordCount > 0 Then
            Dim bano_ant As String
            Dim BANO As String
            Do
                If bano_ant = rs(3) Then
                    BANO = ""
                Else
                    BANO = rs(3)
                    bano_ant = rs(3)
                End If
                
                With lista.ListItems.Add(, , BANO)
                    .SubItems(1) = rs(4)
                    .SubItems(2) = rs(5)
                    .SubItems(3) = rs(6)
                End With
                rs.MoveNext
            Loop Until rs.EOF
        End If
    End If
End Sub

Public Function validar() As Boolean
    If cmbClientes.getTEXTO = "" Then
        MsgBox "Introduzca un cliente para la oferta.", vbCritical, App.Title
        validar = False
        Exit Function
    End If
    If lista.ListItems.Count = 0 Then
        MsgBox "Introduzca algún concepto en la oferta.", vbCritical, App.Title
        validar = False
        Exit Function
    End If
    If opTipo(3).value = True Then
        If opsubTipo(4).value = False And _
            opsubTipo(1).value = False And _
            opsubTipo(2).value = False And _
            opsubTipo(3).value = False Then
            MsgBox "Introduzca el subtipo de oferta.", vbCritical, App.Title
            validar = False
            Exit Function
        End If
    End If
            
    validar = True
End Function

Private Sub opTipo_Click(Index As Integer)
    lista.ListItems.Clear
    tipo = Index
    datos(3).Visible = True
    frameSubtipo.Enabled = False
    frameSubtipo.Visible = False
    Select Case Index
    Case 0, 4 ' General, Agroalimentario
        With lista.ColumnHeaders
            .Clear
            .Add , , "Producto", 4800, lvwColumnLeft
            .Add , , "Parametros", 4700, lvwColumnLeft
            .Add , , "Procedimiento", 1800, lvwColumnLeft
            .Add , , "Precio", 1300, lvwColumnRight
        End With
    Case 1 ' Solucion
        With lista.ColumnHeaders
            .Clear
            .Add , , "Tratamiento-Baño", 4800, lvwColumnLeft
            .Add , , "Determinación", 4700, lvwColumnLeft
            .Add , , "Rango", 1800, lvwColumnLeft
            .Add , , "Precio", 1300, lvwColumnRight
        End With
    Case 2 ' CE
        With lista.ColumnHeaders
            .Clear
            .Add , , "Ensayo", 4800, lvwColumnLeft
            .Add , , "Norma", 4700, lvwColumnLeft
            .Add , , "--", 1800, lvwColumnLeft
            .Add , , "Precio", 1300, lvwColumnRight
        End With
        datos(3).Visible = False
    Case 3 ' Suministro
        With lista.ColumnHeaders
            .Clear
            .Add , , "Concepto", 4800, lvwColumnLeft
            .Add , , "Envase/Unidad", 4700, lvwColumnLeft
            .Add , , "--", 1800, lvwColumnLeft
            .Add , , "Precio", 1300, lvwColumnRight
        End With
        datos(3).Visible = False
        frameSubtipo.Enabled = True
        frameSubtipo.Visible = True
    End Select
End Sub

Public Sub borrar_campos()
    Dim i As Integer
    For i = 1 To 4
        datos(i) = ""
    Next
    calcular_total
End Sub

Public Sub calcular_total()
    Dim i As Integer
    Dim total As Currency
    total = 0
    For i = 1 To lista.ListItems.Count
         If IsNumeric(lista.ListItems(i).SubItems(3)) Then
            total = total + CCur(lista.ListItems(i).SubItems(3))
         End If
    Next
    txtDatos(2) = Format(total, "currency")
End Sub

'Public Property Get NUEVA_EDICION() As Boolean
'    NUEVA_EDICION = mvarblnResultado
'End Property

'Public Property Let NUEVA_EDICION(ByVal blnResultado As Boolean)
'    mvarblnResultado = blnResultado
'End Property

