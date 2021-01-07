VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmUsuarios 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Usuarios y permisos"
   ClientHeight    =   9525
   ClientLeft      =   2430
   ClientTop       =   1380
   ClientWidth     =   13605
   Icon            =   "frmUsuarios.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9525
   ScaleWidth      =   13605
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame frmFNMT 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Revisión de Informes y Firma digital (FNMT)"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1920
      Left            =   45
      TabIndex        =   71
      Top             =   5085
      Visible         =   0   'False
      Width           =   7740
      Begin VB.TextBox datos 
         Appearance      =   0  'Flat
         DataSource      =   "Adodc1"
         Height          =   330
         IMEMode         =   3  'DISABLE
         Index           =   9
         Left            =   1125
         TabIndex        =   10
         Top             =   630
         Width           =   5490
      End
      Begin VB.CheckBox Check1 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Caption         =   "REVISION DE MUESTRAS"
         DataField       =   "PER_ELIMINACION"
         DataSource      =   "Adodc1"
         Enabled         =   0   'False
         ForeColor       =   &H80000008&
         Height          =   375
         Index           =   22
         Left            =   135
         TabIndex        =   9
         Top             =   225
         Width           =   2580
      End
      Begin VB.TextBox datos 
         Appearance      =   0  'Flat
         DataField       =   "PASSWORD"
         DataSource      =   "Adodc1"
         Height          =   330
         IMEMode         =   3  'DISABLE
         Index           =   8
         Left            =   1125
         PasswordChar    =   "*"
         TabIndex        =   13
         Top             =   1395
         Width           =   2535
      End
      Begin VB.CommandButton Command2 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Explorar"
         Height          =   375
         Left            =   6705
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   990
         Width           =   915
      End
      Begin VB.TextBox datos 
         Appearance      =   0  'Flat
         DataSource      =   "Adodc1"
         Height          =   330
         IMEMode         =   3  'DISABLE
         Index           =   7
         Left            =   1125
         TabIndex        =   11
         Top             =   1020
         Width           =   5490
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Cargo Informe"
         Height          =   195
         Index           =   9
         Left            =   90
         TabIndex        =   75
         Top             =   690
         Width           =   990
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Contraseña"
         Height          =   195
         Index           =   8
         Left            =   75
         TabIndex        =   73
         Top             =   1455
         Width           =   825
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Ruta Servidor"
         Height          =   195
         Index           =   7
         Left            =   75
         TabIndex        =   72
         Top             =   1080
         Width           =   1005
      End
   End
   Begin TabDlg.SSTab tabPermisosResponsables 
      Height          =   7860
      Left            =   7875
      TabIndex        =   25
      Top             =   675
      Width           =   5685
      _ExtentX        =   10028
      _ExtentY        =   13864
      _Version        =   393216
      Style           =   1
      TabHeight       =   520
      BackColor       =   12632256
      TabCaption(0)   =   "Propiedades"
      TabPicture(0)   =   "frmUsuarios.frx":27A2
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Check1(14)"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Check1(12)"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Check1(10)"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Check1(9)"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Check1(8)"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Check1(7)"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "Check1(6)"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "Check1(5)"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "Check1(4)"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "Check1(3)"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "Check1(2)"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "Check1(0)"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "Check1(17)"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "Check1(18)"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "Check1(19)"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "Frame1"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).Control(16)=   "Check1(21)"
      Tab(0).Control(16).Enabled=   0   'False
      Tab(0).Control(17)=   "Check1(23)"
      Tab(0).Control(17).Enabled=   0   'False
      Tab(0).Control(18)=   "Frame4"
      Tab(0).Control(18).Enabled=   0   'False
      Tab(0).Control(19)=   "Check1(26)"
      Tab(0).Control(19).Enabled=   0   'False
      Tab(0).Control(20)=   "Frame5"
      Tab(0).Control(20).Enabled=   0   'False
      Tab(0).Control(21)=   "Check1(31)"
      Tab(0).Control(21).Enabled=   0   'False
      Tab(0).Control(22)=   "Frame6"
      Tab(0).Control(22).Enabled=   0   'False
      Tab(0).Control(23)=   "Check1(35)"
      Tab(0).Control(23).Enabled=   0   'False
      Tab(0).Control(24)=   "Check1(1)"
      Tab(0).Control(24).Enabled=   0   'False
      Tab(0).Control(25)=   "Check1(36)"
      Tab(0).Control(25).Enabled=   0   'False
      Tab(0).ControlCount=   26
      TabCaption(1)   =   "Responsables Dpto"
      TabPicture(1)   =   "frmUsuarios.frx":27BE
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "chkDepartamento(10)"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "chkDepartamento(9)"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "chkDepartamento(1)"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "chkDepartamento(2)"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).Control(4)=   "chkDepartamento(3)"
      Tab(1).Control(4).Enabled=   0   'False
      Tab(1).Control(5)=   "chkDepartamento(4)"
      Tab(1).Control(5).Enabled=   0   'False
      Tab(1).Control(6)=   "chkDepartamento(5)"
      Tab(1).Control(6).Enabled=   0   'False
      Tab(1).Control(7)=   "chkDepartamento(6)"
      Tab(1).Control(7).Enabled=   0   'False
      Tab(1).Control(8)=   "chkDepartamento(7)"
      Tab(1).Control(8).Enabled=   0   'False
      Tab(1).Control(9)=   "chkDepartamento(8)"
      Tab(1).Control(9).Enabled=   0   'False
      Tab(1).Control(10)=   "Label2"
      Tab(1).Control(10).Enabled=   0   'False
      Tab(1).ControlCount=   11
      TabCaption(2)   =   "Departamentos"
      TabPicture(2)   =   "frmUsuarios.frx":27DA
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "chkDepartamentoUsuario(8)"
      Tab(2).Control(1)=   "chkDepartamentoUsuario(7)"
      Tab(2).Control(2)=   "chkDepartamentoUsuario(6)"
      Tab(2).Control(3)=   "chkDepartamentoUsuario(5)"
      Tab(2).Control(4)=   "chkDepartamentoUsuario(4)"
      Tab(2).Control(5)=   "chkDepartamentoUsuario(3)"
      Tab(2).Control(6)=   "chkDepartamentoUsuario(2)"
      Tab(2).Control(7)=   "chkDepartamentoUsuario(1)"
      Tab(2).Control(8)=   "chkDepartamentoUsuario(9)"
      Tab(2).Control(9)=   "chkDepartamentoUsuario(10)"
      Tab(2).Control(10)=   "Label3"
      Tab(2).ControlCount=   11
      Begin VB.CheckBox Check1 
         Appearance      =   0  'Flat
         Caption         =   "Mod/Eliminación de Clientes"
         DataSource      =   "Adodc1"
         ForeColor       =   &H80000008&
         Height          =   345
         Index           =   36
         Left            =   90
         TabIndex        =   92
         Top             =   3285
         Width           =   2580
      End
      Begin VB.CheckBox Check1 
         Appearance      =   0  'Flat
         Caption         =   "Facturacion"
         DataField       =   "PER_FACTURACION"
         DataSource      =   "Adodc1"
         ForeColor       =   &H80000008&
         Height          =   345
         Index           =   1
         Left            =   90
         TabIndex        =   37
         Top             =   585
         Width           =   1500
      End
      Begin VB.CheckBox Check1 
         Appearance      =   0  'Flat
         Caption         =   "ENAC: Des. Producto"
         DataSource      =   "Adodc1"
         ForeColor       =   &H80000008&
         Height          =   345
         Index           =   35
         Left            =   90
         TabIndex        =   91
         Top             =   3015
         Width           =   2580
      End
      Begin VB.Frame Frame6 
         Caption         =   "Tesorería"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2130
         Left            =   2925
         TabIndex        =   87
         Top             =   5625
         Width           =   2625
         Begin VB.CheckBox Check1 
            Appearance      =   0  'Flat
            Caption         =   "Facturas Proveedor"
            DataField       =   "PER_ELIMINACION"
            DataSource      =   "Adodc1"
            ForeColor       =   &H80000008&
            Height          =   345
            Index           =   33
            Left            =   90
            TabIndex        =   89
            Top             =   540
            Width           =   2175
         End
         Begin VB.CheckBox Check1 
            Appearance      =   0  'Flat
            Caption         =   "Menu Tesoreria"
            DataField       =   "PER_ELIMINACION"
            DataSource      =   "Adodc1"
            ForeColor       =   &H80000008&
            Height          =   345
            Index           =   32
            Left            =   90
            TabIndex        =   88
            Top             =   225
            Width           =   2175
         End
      End
      Begin VB.CheckBox Check1 
         Appearance      =   0  'Flat
         Caption         =   "Subcontratación Genérica"
         DataField       =   "PER_ELIMINACION"
         DataSource      =   "Adodc1"
         ForeColor       =   &H80000008&
         Height          =   345
         Index           =   31
         Left            =   2880
         TabIndex        =   86
         Top             =   2520
         Width           =   2760
      End
      Begin VB.Frame Frame5 
         Caption         =   "Formación"
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
         Left            =   90
         TabIndex        =   83
         Top             =   6750
         Width           =   2805
         Begin VB.CheckBox Check1 
            Appearance      =   0  'Flat
            Caption         =   "Cursos RFI"
            DataField       =   "PER_ELIMINACION"
            DataSource      =   "Adodc1"
            ForeColor       =   &H80000008&
            Height          =   345
            Index           =   29
            Left            =   90
            TabIndex        =   85
            Top             =   225
            Width           =   2175
         End
         Begin VB.CheckBox Check1 
            Appearance      =   0  'Flat
            Caption         =   "Plan Formación Anual"
            DataField       =   "PER_ELIMINACION"
            DataSource      =   "Adodc1"
            ForeColor       =   &H80000008&
            Height          =   345
            Index           =   30
            Left            =   90
            TabIndex        =   84
            Top             =   540
            Width           =   1860
         End
      End
      Begin VB.CheckBox Check1 
         Appearance      =   0  'Flat
         Caption         =   "Tramitación de subcontratación"
         DataField       =   "PER_ELIMINACION"
         DataSource      =   "Adodc1"
         ForeColor       =   &H80000008&
         Height          =   345
         Index           =   26
         Left            =   2865
         TabIndex        =   82
         Top             =   1890
         Width           =   2760
      End
      Begin VB.Frame Frame4 
         Caption         =   "Indicadores"
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
         Left            =   2925
         TabIndex        =   77
         Top             =   3870
         Width           =   2625
         Begin VB.CheckBox Check1 
            Appearance      =   0  'Flat
            Caption         =   "Productividad"
            DataField       =   "PER_MATRIZ_CUALIFICACIONES"
            DataSource      =   "Adodc1"
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   38
            Left            =   90
            TabIndex        =   94
            Top             =   1260
            Width           =   2310
         End
         Begin VB.CheckBox Check1 
            Appearance      =   0  'Flat
            Caption         =   "Cambio de Plazo de Entrega"
            DataField       =   "PER_ELIMINACION"
            DataSource      =   "Adodc1"
            ForeColor       =   &H80000008&
            Height          =   345
            Index           =   27
            Left            =   90
            TabIndex        =   79
            Top             =   270
            Width           =   2355
         End
         Begin VB.CheckBox Check1 
            Appearance      =   0  'Flat
            Caption         =   "Listados de Fuera de Plazo"
            DataField       =   "PER_ELIMINACION"
            DataSource      =   "Adodc1"
            ForeColor       =   &H80000008&
            Height          =   345
            Index           =   25
            Left            =   90
            TabIndex        =   78
            Top             =   585
            Width           =   2445
         End
         Begin VB.CheckBox Check1 
            Appearance      =   0  'Flat
            Caption         =   "Indicadores Cliente"
            DataField       =   "PER_MATRIZ_CUALIFICACIONES"
            DataSource      =   "Adodc1"
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   28
            Left            =   90
            TabIndex        =   80
            Top             =   945
            Width           =   2310
         End
      End
      Begin VB.CheckBox Check1 
         Appearance      =   0  'Flat
         Caption         =   "Ofertas"
         DataField       =   "PER_ELIMINACION"
         DataSource      =   "Adodc1"
         ForeColor       =   &H80000008&
         Height          =   345
         Index           =   23
         Left            =   2865
         TabIndex        =   76
         Top             =   1575
         Width           =   1815
      End
      Begin VB.CheckBox Check1 
         Appearance      =   0  'Flat
         Caption         =   "Gestión de Incurridos"
         DataField       =   "PER_ELIMINACION"
         DataSource      =   "Adodc1"
         ForeColor       =   &H80000008&
         Height          =   345
         Index           =   21
         Left            =   2865
         TabIndex        =   74
         Top             =   1260
         Width           =   1905
      End
      Begin VB.Frame Frame1 
         Caption         =   "Calidad"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2850
         Left            =   45
         TabIndex        =   65
         Top             =   3870
         Width           =   2850
         Begin VB.CheckBox Check1 
            Appearance      =   0  'Flat
            Caption         =   "Listado de Incidencias"
            DataField       =   "PER_ELIMINACION"
            DataSource      =   "Adodc1"
            ForeColor       =   &H80000008&
            Height          =   345
            Index           =   37
            Left            =   90
            TabIndex        =   93
            Top             =   2475
            Width           =   2490
         End
         Begin VB.CheckBox Check1 
            Appearance      =   0  'Flat
            Caption         =   "Normas NO CONTROLADAS"
            DataField       =   "PER_ELIMINACION"
            DataSource      =   "Adodc1"
            ForeColor       =   &H80000008&
            Height          =   345
            Index           =   34
            Left            =   90
            TabIndex        =   90
            Top             =   2160
            Width           =   2490
         End
         Begin VB.CheckBox Check1 
            Appearance      =   0  'Flat
            Caption         =   "Ver todas las Familias"
            DataField       =   "PER_ELIMINACION"
            DataSource      =   "Adodc1"
            ForeColor       =   &H80000008&
            Height          =   345
            Index           =   24
            Left            =   90
            TabIndex        =   81
            Top             =   1530
            Width           =   2175
         End
         Begin VB.CheckBox Check1 
            Appearance      =   0  'Flat
            Caption         =   "Gestión de No Conformidades"
            DataField       =   "PER_ELIMINACION"
            DataSource      =   "Adodc1"
            ForeColor       =   &H80000008&
            Height          =   345
            Index           =   13
            Left            =   90
            TabIndex        =   70
            Top             =   1845
            Width           =   2490
         End
         Begin VB.CheckBox Check1 
            Appearance      =   0  'Flat
            Caption         =   "Creación Versiones Documentos"
            DataField       =   "PER_ELIMINACION"
            DataSource      =   "Adodc1"
            ForeColor       =   &H80000008&
            Height          =   345
            Index           =   15
            Left            =   90
            TabIndex        =   69
            Top             =   585
            Width           =   2625
         End
         Begin VB.CheckBox Check1 
            Appearance      =   0  'Flat
            Caption         =   "Superusuario PNT"
            DataField       =   "PER_ELIMINACION"
            DataSource      =   "Adodc1"
            ForeColor       =   &H80000008&
            Height          =   345
            Index           =   16
            Left            =   90
            TabIndex        =   68
            Top             =   1215
            Width           =   2175
         End
         Begin VB.CheckBox Check1 
            Appearance      =   0  'Flat
            Caption         =   "Acceso Documentación Calidad"
            DataField       =   "PER_ELIMINACION"
            DataSource      =   "Adodc1"
            ForeColor       =   &H80000008&
            Height          =   345
            Index           =   11
            Left            =   90
            TabIndex        =   67
            Top             =   270
            Width           =   2580
         End
         Begin VB.CheckBox Check1 
            Appearance      =   0  'Flat
            Caption         =   "Impresión docs. calidad"
            DataField       =   "PER_MATRIZ_CUALIFICACIONES"
            DataSource      =   "Adodc1"
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   20
            Left            =   90
            TabIndex        =   66
            Top             =   945
            Width           =   2625
         End
      End
      Begin VB.CheckBox Check1 
         Appearance      =   0  'Flat
         Caption         =   "Videos"
         DataField       =   "PER_MATRIZ_CUALIFICACIONES"
         DataSource      =   "Adodc1"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   19
         Left            =   2880
         TabIndex        =   64
         Top             =   2925
         Width           =   2400
      End
      Begin VB.CheckBox Check1 
         Appearance      =   0  'Flat
         Caption         =   "Matriz de Cualificaciones"
         DataField       =   "PER_MATRIZ_CUALIFICACIONES"
         DataSource      =   "Adodc1"
         ForeColor       =   &H80000008&
         Height          =   345
         Index           =   18
         Left            =   90
         TabIndex        =   63
         Top             =   2475
         Width           =   2670
      End
      Begin VB.CheckBox Check1 
         Appearance      =   0  'Flat
         Caption         =   "Datos Especiales Muestras"
         DataField       =   "PER_ELIMINACION"
         DataSource      =   "Adodc1"
         ForeColor       =   &H80000008&
         Height          =   345
         Index           =   17
         Left            =   2880
         TabIndex        =   61
         Top             =   3195
         Width           =   2535
      End
      Begin VB.CheckBox chkDepartamentoUsuario 
         Appearance      =   0  'Flat
         Caption         =   "I + D"
         DataField       =   "PER_ELIMINACION"
         DataSource      =   "Adodc1"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   345
         Index           =   8
         Left            =   -74910
         TabIndex        =   58
         Top             =   2535
         Width           =   3300
      End
      Begin VB.CheckBox chkDepartamentoUsuario 
         Appearance      =   0  'Flat
         Caption         =   "Logística"
         DataField       =   "PER_ELIMINACION"
         DataSource      =   "Adodc1"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   345
         Index           =   7
         Left            =   -74910
         TabIndex        =   57
         Top             =   2220
         Width           =   3300
      End
      Begin VB.CheckBox chkDepartamentoUsuario 
         Appearance      =   0  'Flat
         Caption         =   "Metrología"
         DataField       =   "PER_ELIMINACION"
         DataSource      =   "Adodc1"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   345
         Index           =   6
         Left            =   -74910
         TabIndex        =   56
         Top             =   1905
         Width           =   3300
      End
      Begin VB.CheckBox chkDepartamentoUsuario 
         Appearance      =   0  'Flat
         Caption         =   "Laborat. Aeronáutico"
         DataField       =   "PER_ELIMINACION"
         DataSource      =   "Adodc1"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   345
         Index           =   5
         Left            =   -74910
         TabIndex        =   55
         Top             =   1590
         Width           =   3300
      End
      Begin VB.CheckBox chkDepartamentoUsuario 
         Appearance      =   0  'Flat
         Caption         =   "Laborat. Agroalimentario"
         DataField       =   "PER_ELIMINACION"
         DataSource      =   "Adodc1"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   345
         Index           =   4
         Left            =   -74910
         TabIndex        =   54
         Top             =   1275
         Width           =   3300
      End
      Begin VB.CheckBox chkDepartamentoUsuario 
         Appearance      =   0  'Flat
         Caption         =   "Administación y RRHH"
         DataField       =   "PER_ELIMINACION"
         DataSource      =   "Adodc1"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   345
         Index           =   3
         Left            =   -74910
         TabIndex        =   53
         Top             =   960
         Width           =   3300
      End
      Begin VB.CheckBox chkDepartamentoUsuario 
         Appearance      =   0  'Flat
         Caption         =   "Calidad"
         DataField       =   "PER_ELIMINACION"
         DataSource      =   "Adodc1"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   345
         Index           =   2
         Left            =   -74910
         TabIndex        =   52
         Top             =   645
         Width           =   3165
      End
      Begin VB.CheckBox chkDepartamentoUsuario 
         Appearance      =   0  'Flat
         Caption         =   "Gerencia"
         DataField       =   "PER_ELIMINACION"
         DataSource      =   "Adodc1"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   345
         Index           =   1
         Left            =   -74910
         TabIndex        =   51
         Top             =   330
         Width           =   2760
      End
      Begin VB.CheckBox chkDepartamentoUsuario 
         Appearance      =   0  'Flat
         Caption         =   "Informática"
         DataField       =   "PER_ELIMINACION"
         DataSource      =   "Adodc1"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   345
         Index           =   9
         Left            =   -74910
         TabIndex        =   50
         Top             =   2850
         Width           =   3300
      End
      Begin VB.CheckBox chkDepartamentoUsuario 
         Appearance      =   0  'Flat
         Caption         =   "Recepción"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   345
         Index           =   10
         Left            =   -74910
         TabIndex        =   49
         Top             =   3195
         Width           =   2520
      End
      Begin VB.CheckBox chkDepartamento 
         Appearance      =   0  'Flat
         Caption         =   "Recepción"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   345
         Index           =   10
         Left            =   -74910
         TabIndex        =   48
         Top             =   3195
         Width           =   2520
      End
      Begin VB.CheckBox chkDepartamento 
         Appearance      =   0  'Flat
         Caption         =   "Informática"
         DataField       =   "PER_ELIMINACION"
         DataSource      =   "Adodc1"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   345
         Index           =   9
         Left            =   -74910
         TabIndex        =   47
         Top             =   2850
         Width           =   3300
      End
      Begin VB.CheckBox chkDepartamento 
         Appearance      =   0  'Flat
         Caption         =   "Gerencia"
         DataField       =   "PER_ELIMINACION"
         DataSource      =   "Adodc1"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   345
         Index           =   1
         Left            =   -74910
         TabIndex        =   46
         Top             =   330
         Width           =   2760
      End
      Begin VB.CheckBox chkDepartamento 
         Appearance      =   0  'Flat
         Caption         =   "Calidad"
         DataField       =   "PER_ELIMINACION"
         DataSource      =   "Adodc1"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   345
         Index           =   2
         Left            =   -74910
         TabIndex        =   45
         Top             =   645
         Width           =   3165
      End
      Begin VB.CheckBox chkDepartamento 
         Appearance      =   0  'Flat
         Caption         =   "Administación y RRHH"
         DataField       =   "PER_ELIMINACION"
         DataSource      =   "Adodc1"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   345
         Index           =   3
         Left            =   -74910
         TabIndex        =   44
         Top             =   960
         Width           =   3300
      End
      Begin VB.CheckBox chkDepartamento 
         Appearance      =   0  'Flat
         Caption         =   "Laborat. Agroalimentario"
         DataField       =   "PER_ELIMINACION"
         DataSource      =   "Adodc1"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   345
         Index           =   4
         Left            =   -74910
         TabIndex        =   43
         Top             =   1275
         Width           =   3300
      End
      Begin VB.CheckBox chkDepartamento 
         Appearance      =   0  'Flat
         Caption         =   "Laborat. Aeronáutico"
         DataField       =   "PER_ELIMINACION"
         DataSource      =   "Adodc1"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   345
         Index           =   5
         Left            =   -74910
         TabIndex        =   42
         Top             =   1590
         Width           =   3300
      End
      Begin VB.CheckBox chkDepartamento 
         Appearance      =   0  'Flat
         Caption         =   "Metrología"
         DataField       =   "PER_ELIMINACION"
         DataSource      =   "Adodc1"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   345
         Index           =   6
         Left            =   -74910
         TabIndex        =   41
         Top             =   1905
         Width           =   3300
      End
      Begin VB.CheckBox chkDepartamento 
         Appearance      =   0  'Flat
         Caption         =   "Logística"
         DataField       =   "PER_ELIMINACION"
         DataSource      =   "Adodc1"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   345
         Index           =   7
         Left            =   -74910
         TabIndex        =   40
         Top             =   2220
         Width           =   3300
      End
      Begin VB.CheckBox chkDepartamento 
         Appearance      =   0  'Flat
         Caption         =   "I + D"
         DataField       =   "PER_ELIMINACION"
         DataSource      =   "Adodc1"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   345
         Index           =   8
         Left            =   -74910
         TabIndex        =   39
         Top             =   2535
         Width           =   3300
      End
      Begin VB.CheckBox Check1 
         Appearance      =   0  'Flat
         Caption         =   "Impresion"
         DataField       =   "PER_IMPRESION"
         DataSource      =   "Adodc1"
         ForeColor       =   &H80000008&
         Height          =   300
         Index           =   0
         Left            =   90
         TabIndex        =   38
         Top             =   330
         Width           =   1365
      End
      Begin VB.CheckBox Check1 
         Appearance      =   0  'Flat
         Caption         =   "Modificacion"
         DataField       =   "PER_MODIFICACION"
         DataSource      =   "Adodc1"
         ForeColor       =   &H80000008&
         Height          =   345
         Index           =   2
         Left            =   90
         TabIndex        =   36
         Top             =   855
         Width           =   1725
      End
      Begin VB.CheckBox Check1 
         Appearance      =   0  'Flat
         Caption         =   "Eliminacion"
         DataField       =   "PER_MODIFICACION"
         DataSource      =   "Adodc1"
         ForeColor       =   &H80000008&
         Height          =   345
         Index           =   3
         Left            =   90
         TabIndex        =   35
         Top             =   1125
         Width           =   1635
      End
      Begin VB.CheckBox Check1 
         Appearance      =   0  'Flat
         Caption         =   "Altas / Bajas usuarios"
         DataField       =   "PER_ELIMINACION"
         DataSource      =   "Adodc1"
         ForeColor       =   &H80000008&
         Height          =   345
         Index           =   4
         Left            =   90
         TabIndex        =   34
         Top             =   1395
         Width           =   2580
      End
      Begin VB.CheckBox Check1 
         Appearance      =   0  'Flat
         Caption         =   "Segunda Edición"
         DataField       =   "PER_ELIMINACION"
         DataSource      =   "Adodc1"
         ForeColor       =   &H80000008&
         Height          =   345
         Index           =   5
         Left            =   90
         TabIndex        =   33
         Top             =   1665
         Width           =   2580
      End
      Begin VB.CheckBox Check1 
         Appearance      =   0  'Flat
         Caption         =   "Cierre de Muestras"
         DataField       =   "PER_ELIMINACION"
         DataSource      =   "Adodc1"
         ForeColor       =   &H80000008&
         Height          =   345
         Index           =   6
         Left            =   90
         TabIndex        =   32
         Top             =   1935
         Width           =   2580
      End
      Begin VB.CheckBox Check1 
         Appearance      =   0  'Flat
         Caption         =   "Pedidos de Reactivos"
         DataField       =   "PER_ELIMINACION"
         DataSource      =   "Adodc1"
         ForeColor       =   &H80000008&
         Height          =   345
         Index           =   7
         Left            =   90
         TabIndex        =   31
         Top             =   2205
         Width           =   2580
      End
      Begin VB.CheckBox Check1 
         Appearance      =   0  'Flat
         Caption         =   "Gestión de Empleados"
         DataField       =   "PER_ELIMINACION"
         DataSource      =   "Adodc1"
         ForeColor       =   &H80000008&
         Height          =   345
         Index           =   8
         Left            =   2865
         TabIndex        =   30
         Top             =   2205
         Visible         =   0   'False
         Width           =   2310
      End
      Begin VB.CheckBox Check1 
         Appearance      =   0  'Flat
         Caption         =   "Contabilidad"
         DataField       =   "PER_ELIMINACION"
         DataSource      =   "Adodc1"
         ForeColor       =   &H80000008&
         Height          =   315
         Index           =   9
         Left            =   2865
         TabIndex        =   29
         Top             =   360
         Width           =   1455
      End
      Begin VB.CheckBox Check1 
         Appearance      =   0  'Flat
         Caption         =   "Envío Pedidos a Proveedor"
         DataField       =   "PER_ELIMINACION"
         DataSource      =   "Adodc1"
         ForeColor       =   &H80000008&
         Height          =   315
         Index           =   10
         Left            =   2865
         TabIndex        =   28
         Top             =   645
         Width           =   2400
      End
      Begin VB.CheckBox Check1 
         Appearance      =   0  'Flat
         Caption         =   "Gestión de proyectos"
         DataField       =   "PER_ELIMINACION"
         DataSource      =   "Adodc1"
         ForeColor       =   &H80000008&
         Height          =   345
         Index           =   12
         Left            =   2865
         TabIndex        =   27
         Top             =   945
         Width           =   1905
      End
      Begin VB.CheckBox Check1 
         Appearance      =   0  'Flat
         Caption         =   "Tareas"
         DataField       =   "PER_ELIMINACION"
         DataSource      =   "Adodc1"
         ForeColor       =   &H80000008&
         Height          =   345
         Index           =   14
         Left            =   90
         TabIndex        =   26
         Top             =   2745
         Width           =   2580
      End
      Begin VB.Label Label3 
         Caption         =   "Señalar ÚNICAMENTE aquellos departamentos de los cuales PERTENECE el usuario"
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
         Height          =   435
         Left            =   -74910
         TabIndex        =   60
         Top             =   4140
         Width           =   5535
      End
      Begin VB.Label Label2 
         Caption         =   "Señalar ÚNICAMENTE aquellos departamentos de los cuales el usuario figure como Responsable"
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
         Height          =   480
         Left            =   -74910
         TabIndex        =   59
         Top             =   4050
         Width           =   5535
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Firma electrónica"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   45
      TabIndex        =   21
      Top             =   3690
      Width           =   7740
      Begin VB.CommandButton cmdEXplorar 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Explorar"
         Height          =   375
         Left            =   4815
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   585
         Width           =   915
      End
      Begin VB.TextBox datos 
         Appearance      =   0  'Flat
         DataSource      =   "Adodc1"
         Height          =   330
         IMEMode         =   3  'DISABLE
         Index           =   4
         Left            =   720
         TabIndex        =   7
         Top             =   615
         Width           =   4050
      End
      Begin VB.Image img 
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
         Height          =   945
         Left            =   5850
         Stretch         =   -1  'True
         Top             =   270
         Width           =   1815
      End
      Begin VB.Label lblCampos 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Ruta"
         Height          =   195
         Index           =   4
         Left            =   165
         TabIndex        =   22
         Top             =   675
         Width           =   375
      End
   End
   Begin VB.CommandButton cmdcancel 
      BackColor       =   &H00E0E0E0&
      Caption         =   "ESC-Salir"
      Height          =   870
      Left            =   12495
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   8595
      Width           =   1050
   End
   Begin VB.CommandButton cmdok 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Aceptar"
      Height          =   870
      Left            =   11370
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   8595
      Width           =   1050
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Datos Personales"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3090
      Left            =   45
      TabIndex        =   16
      Top             =   585
      Width           =   7755
      Begin VB.TextBox datos 
         Appearance      =   0  'Flat
         DataField       =   "PASSWORD"
         DataSource      =   "Adodc1"
         Height          =   330
         IMEMode         =   3  'DISABLE
         Index           =   6
         Left            =   1035
         TabIndex        =   6
         Top             =   2430
         Width           =   5010
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Explorar"
         Height          =   375
         Left            =   6210
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   2025
         Width           =   1320
      End
      Begin VB.TextBox datos 
         Appearance      =   0  'Flat
         DataField       =   "PASSWORD"
         DataSource      =   "Adodc1"
         Height          =   330
         IMEMode         =   3  'DISABLE
         Index           =   5
         Left            =   1035
         TabIndex        =   4
         Top             =   2025
         Width           =   5010
      End
      Begin VB.TextBox datos 
         Appearance      =   0  'Flat
         DataField       =   "PASSWORD"
         DataSource      =   "Adodc1"
         Height          =   330
         IMEMode         =   3  'DISABLE
         Index           =   3
         Left            =   1035
         PasswordChar    =   "*"
         TabIndex        =   3
         Top             =   1620
         Width           =   2535
      End
      Begin VB.TextBox datos 
         Appearance      =   0  'Flat
         DataField       =   "NOMBRE"
         DataSource      =   "Adodc1"
         Height          =   360
         Index           =   1
         Left            =   1035
         TabIndex        =   1
         Top             =   810
         Width           =   5070
      End
      Begin VB.TextBox datos 
         Appearance      =   0  'Flat
         DataField       =   "APELLIDOS"
         DataSource      =   "Adodc1"
         Height          =   330
         Index           =   2
         Left            =   1035
         TabIndex        =   2
         Top             =   1215
         Width           =   5070
      End
      Begin VB.TextBox datos 
         Appearance      =   0  'Flat
         DataField       =   "USUARIO"
         DataSource      =   "Adodc1"
         Height          =   360
         Index           =   0
         Left            =   1035
         TabIndex        =   0
         Top             =   405
         Width           =   5070
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "E-Mail"
         Height          =   195
         Index           =   6
         Left            =   90
         TabIndex        =   24
         Top             =   2475
         Width           =   465
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Imagen"
         Height          =   195
         Index           =   5
         Left            =   90
         TabIndex        =   23
         Top             =   2070
         Width           =   555
      End
      Begin VB.Image Image3 
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
         Height          =   1710
         Left            =   6165
         Stretch         =   -1  'True
         Top             =   225
         Width           =   1470
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Nombre"
         Height          =   195
         Index           =   2
         Left            =   90
         TabIndex        =   20
         Top             =   855
         Width           =   585
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Apellidos"
         Height          =   195
         Index           =   0
         Left            =   90
         TabIndex        =   19
         Top             =   1260
         Width           =   645
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Usuario"
         Height          =   195
         Index           =   1
         Left            =   90
         TabIndex        =   18
         Top             =   450
         Width           =   555
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Contraseña"
         Height          =   195
         Index           =   3
         Left            =   90
         TabIndex        =   17
         Top             =   1680
         Width           =   825
      End
   End
   Begin MSComDlg.CommonDialog cd 
      Left            =   6450
      Top             =   5655
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   13050
      Picture         =   "frmUsuarios.frx":27F6
      Top             =   0
      Width           =   480
   End
   Begin VB.Image imagen 
      Height          =   480
      Left            =   19665
      Picture         =   "frmUsuarios.frx":30C0
      Top             =   -3375
      Width           =   480
   End
   Begin VB.Label lbltitulo 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "Alta de Usuario Nuevo"
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
      TabIndex        =   62
      Top             =   135
      Width           =   2355
   End
   Begin VB.Shape fondo 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   555
      Left            =   0
      Top             =   0
      Width           =   13740
   End
End
Attribute VB_Name = "frmUsuarios"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public PK As Long
Private Const TOTAL_DEPARTAMENTOS = 10
Private mvarstrResposabilidadesDepartamento(TOTAL_DEPARTAMENTOS) As String
Private mvarobjDepartamentosUsuario As clsGenericCollection


Private Sub cmdcancel_Click()
    Unload Me
End Sub

Private Sub cmdok_Click()
    If Trim(datos(0)) = "" Then
        MsgBox "Indique el usuario.", vbExclamation, App.Title
        datos(0).SetFocus
        Exit Sub
    End If
    If Trim(datos(1)) = "" Then
        MsgBox "Indique el nombre.", vbExclamation, App.Title
        datos(1).SetFocus
        Exit Sub
    End If
    If Trim(datos(2)) = "" Then
        MsgBox "Indique el apellido.", vbExclamation, App.Title
        datos(2).SetFocus
        Exit Sub
    End If
    If Trim(datos(3)) = "" Then
        MsgBox "Indique la password.", vbExclamation, App.Title
        datos(3).SetFocus
        Exit Sub
    End If
    Dim oUsu As New clsUsuarios
    Dim intCont As Integer
    Dim objdep As clsGenericClass
    
    With oUsu
            .setNOMBRE = datos(1)
            .setAPELLIDOS = datos(2)
            .setUSUARIO = datos(0)
            .setPASSWORD = Encripta(datos(3), datos(0))
            .setREVISION_CARGO = datos(9)
            ' Per. Impresion
            If Check1(0).Value = Checked Then
                .setPER_IMPRESION = True
            Else
                .setPER_IMPRESION = False
            End If
            ' Per. Facturacion
            If Check1(1).Value = Checked Then
                .setPER_FACTURACION = True
            Else
                .setPER_FACTURACION = False
            End If
            ' Per. Modificacion
            If Check1(2).Value = Checked Then
                .setPER_MODIFICACION = True
            Else
                .setPER_MODIFICACION = False
            End If
            ' Per. Eliminacion
            If Check1(3).Value = Checked Then
                .setPER_ELIMINACION = True
            Else
                .setPER_ELIMINACION = False
            End If
            ' Per. Usuario
            If Check1(4).Value = Checked Then
                .setPER_USUARIOS = True
            Else
                .setPER_USUARIOS = False
            End If
            ' Per. Segunda edicion
            If Check1(5).Value = Checked Then
                .setPER_EDICION = True
            Else
                .setPER_EDICION = False
            End If
            ' Per. Cierre
            If Check1(6).Value = Checked Then
                .setPER_CIERRE = True
            Else
                .setPER_CIERRE = False
            End If
            ' Per. Pedidos de Reactivos
            If Check1(7).Value = Checked Then
                .setPER_PEDIDOS_REACTIVOS = True
            Else
                .setPER_PEDIDOS_REACTIVOS = False
            End If
            ' Per. Empleados
            If Check1(8).Value = Checked Then
                .setPER_EMPLEADOS = True
            Else
                .setPER_EMPLEADOS = False
            End If
            ' Per. Contabilidad
            If Check1(9).Value = Checked Then
                .setPER_CONTABILIDAD = True
            Else
                .setPER_CONTABILIDAD = False
            End If
            ' Per. Pedidos proveedor
            If Check1(10).Value = Checked Then
                .setPER_ENVIO_PEDIDOS_PROVEEDOR = True
            Else
                .setPER_ENVIO_PEDIDOS_PROVEEDOR = False
            End If
            ' Per. Calidad
            If Check1(11).Value = Checked Then
                .setPER_DOCUMENTACION_CALIDAD = True
            Else
                .setPER_DOCUMENTACION_CALIDAD = False
            End If
            ' Per. Proyectos
            If Check1(12).Value = Checked Then
                .setPER_PROYECTOS = True
            Else
                .setPER_PROYECTOS = False
            End If
            ' Per. NC (Gestión de incidencias)
            If Check1(13).Value = Checked Then
                .setPER_NC = True
            Else
                .setPER_NC = False
            End If
            ' Per. PNT
            If Check1(15).Value = Checked Then
                .setPER_PNT = True
            Else
                .setPER_PNT = False
            End If
            If Check1(16).Value = Checked Then
                .setPER_ADMIN_PNT = True
            Else
                .setPER_ADMIN_PNT = False
            End If
            If Check1(17).Value = Checked Then
                .setPER_DATOS_ESPECIALES = True
            Else
                .setPER_DATOS_ESPECIALES = False
            End If
            ' Per. Matriz Cualificaciones
            If Check1(18).Value = Checked Then
                .setPER_MATRIZ_CUALIF = True
            Else
                .setPER_MATRIZ_CUALIF = False
            End If
            If Check1(19).Value = Checked Then
                .setPER_VIDEOS = True
            Else
                .setPER_VIDEOS = False
            End If
            ' Per. Impresion PNT
            If Check1(20).Value = Checked Then
                .setPER_IMPRESION_PNT = True
            Else
                .setPER_IMPRESION_PNT = False
            End If
            ' Per. Incurridos
            If Check1(21).Value = Checked Then
                .setPER_INCURRIDOS = True
            Else
                .setPER_INCURRIDOS = False
            End If
            ' Per. Revision
            If Check1(22).Value = Checked Then
                .setPER_REVISION = True
            Else
                .setPER_REVISION = False
            End If
            ' Per. Ofertas
            If Check1(23).Value = Checked Then
                .setPER_OFERTAS = True
            Else
                .setPER_OFERTAS = False
            End If
            If Check1(27).Value = Checked Then
                .setPER_PLAZO_ENTREGA_CAMBIO = True
            Else
                .setPER_PLAZO_ENTREGA_CAMBIO = False
            End If
            If Check1(25).Value = Checked Then
                .setPER_PLAZO_ENTREGA_LISTADO = True
            Else
                .setPER_PLAZO_ENTREGA_LISTADO = False
            End If
            If Check1(28).Value = Checked Then
                .setPER_INDICADORES_CLIENTE = True
            Else
                .setPER_INDICADORES_CLIENTE = False
            End If
            If Check1(24).Value = Checked Then
                .setPER_FAMILIAS_CA = True
            Else
                .setPER_FAMILIAS_CA = False
            End If
            'M1144-I
            If Check1(26).Value = Checked Then
                .setPER_TRAMITACION_CONTRATA = True
            Else
                .setPER_TRAMITACION_CONTRATA = False
            End If
            If Check1(29).Value = Checked Then
                .setPER_RFI = True
            Else
                .setPER_RFI = False
            End If
            If Check1(30).Value = Checked Then
                .setPER_PFA = True
            Else
                .setPER_PFA = False
            End If
            'M1144-F
            If Check1(31).Value = Checked Then
                .setPER_SCG = True
            Else
                .setPER_SCG = False
            End If
            If Check1(32).Value = Checked Then
                .setPER_TESORERIA_MENU = True
            Else
                .setPER_TESORERIA_MENU = False
            End If
            If Check1(33).Value = Checked Then
                .setPER_TESORERIA_FP = True
            Else
                .setPER_TESORERIA_FP = False
            End If
            'M1377-I
            If Check1(34).Value = Checked Then
                .setPER_NORMAS_NO_CONTROLADAS = True
            Else
                .setPER_NORMAS_NO_CONTROLADAS = False
            End If
            If Check1(35).Value = Checked Then
                .setPER_DES_PRODUCTO = True
            Else
                .setPER_DES_PRODUCTO = False
            End If
            If Check1(36).Value = Checked Then
                .setPER_MOD_CLIENTE = True
            Else
                .setPER_MOD_CLIENTE = False
            End If
            If Check1(37).Value = Checked Then
                .setPER_INCIDENCIAS = True
            Else
                .setPER_INCIDENCIAS = False
            End If
            If Check1(38).Value = Checked Then
                .setPER_PRODUCTIVIDAD = True
            Else
                .setPER_PRODUCTIVIDAD = False
            End If
            
            'M1377-F
            'JONATHAN.2009.10.27
                For intCont = 1 To TOTAL_DEPARTAMENTOS
                    .setRESPONSABLE_DEPARTAMENTOS(intCont) = CInt(chkDepartamento(intCont).Value)
                    .setRESPONSABLE_DEPARTAMENTOS_INICIAL(intCont) = mvarstrResposabilidadesDepartamento(intCont)
                
                If mvarobjDepartamentosUsuario Is Nothing Then
                    Set mvarobjDepartamentosUsuario = New clsGenericCollection
                End If
                
                    Set objdep = mvarobjDepartamentosUsuario.Item(CStr(intCont))
                    If Not objdep Is Nothing Then
                        If chkDepartamentoUsuario(intCont).Value = vbUnchecked Then
                            objdep.setID_AUX = enumIdAux.ID_AUX_ELIMINADO
                            Call mvarobjDepartamentosUsuario.Replace(objdep.getID, objdep)
                        End If
                    Else
                        If chkDepartamentoUsuario(intCont).Value = vbChecked Then
                            Set objdep = New clsGenericClass
                            objdep.setID = intCont
                            objdep.setDESCRIPCION = chkDepartamentoUsuario(intCont).Caption
                            Call mvarobjDepartamentosUsuario.Add(objdep, objdep.getID, enumIdAux.ID_AUX_NUEVO)
                        End If
                    End If
                                    
                Next intCont
                
                Set .setDEPARTAMENTOS_USUARIO = mvarobjDepartamentosUsuario
            
            'FIN JONATHAN
            
            
            .setFIRMA = Replace(datos(4).Text, "\", "/")
            .setIMAGEN = Replace(datos(5).Text, "\", "/")
            .setANULADO = 0
            'E0143-I
            .setEMAIL = datos(6)
            .setFNMT_RUTA = Replace(datos(7), "\", "/")
            .setFNMT_PASS = datos(8)
            'E0143-F
            If gempleado = 0 Then ' Nuevo
                gempleado = .Insertar
                If gempleado <> 0 Then
                    MsgBox "El usuario se ha insertado correctamente", vbInformation, App.Title
                    Unload Me
                End If
            Else
                If .Modificar(gempleado) <> 0 Then
                    MsgBox "El usuario se ha modificado correctamente", vbInformation, App.Title
                    Unload Me
                End If
            End If
    End With
End Sub
   
Private Sub Command1_Click()
    cd.DialogTitle = "Abrir fichero de imagen"
    cd.InitDir = ReadINI(App.Path & "\config.ini", "documentos", "firmas")
    cd.ShowOpen
    If cd.FileName <> "" Then
        datos(5).Text = cd.FileName ' cd.FileTitle
    End If
End Sub

Private Sub Command2_Click()
    cd.DialogTitle = "Abrir fichero de firma"
    cd.InitDir = ReadINI(App.Path & "\config.ini", "documentos", "firmas")
    cd.ShowOpen
    If cd.FileName <> "" Then
        datos(7).Text = cd.FileName ' cd.FileTitle
    End If

End Sub

Private Sub Form_Activate()
    If USUARIO.getPER_USUARIOS = False Then
        MsgBox "No tiene permisos para ver los usuarios.", vbExclamation, App.Title
        Unload Me
        Exit Sub
    End If
    permisos
    If PK <> 0 Then
        gempleado = PK
    End If
    If gempleado <> 0 Then
        cargar_usuario
    Else
        Dim x As Integer
        For x = 1 To 10
            mvarstrResposabilidadesDepartamento(x) = "0"
        Next x
    End If
    If UCase(USUARIO.getUSUARIO) = "JULIO" Then
        datos(3).PasswordChar = ""
        frmFNMT.visible = True
    End If

End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = 27 Then
        cmdcancel_Click
    End If
End Sub
Private Sub Form_Load()
    log (Me.Name)
    cargar_botones Me
End Sub
Public Sub cargar_usuario()
    Dim clsusu As New clsUsuarios
    Dim objDept As clsGenericClass
    Dim intCont As Integer
    
    If clsusu.CARGAR(gempleado) = True Then
     With clsusu
        datos(0) = .getUSUARIO
        datos(1) = .getNOMBRE
        datos(2) = .getAPELLIDOS
        datos(3) = Desencripta(.getPASSWORD, .getUSUARIO)
        datos(9) = .getREVISION_CARGO
'        datos(3) = .getPASSWORD
        If .getPER_IMPRESION = True Then
            Check1(0).Value = Checked
        End If
        If .getPER_FACTURACION = True Then
            Check1(1).Value = Checked
        End If
        If .getPER_MODIFICACION = True Then
            Check1(2).Value = Checked
        End If
        If .getPER_ELIMINACION = True Then
            Check1(3).Value = Checked
        End If
        If .getPER_USUARIOS = True Then
            Check1(4).Value = Checked
        End If
        If .getPER_EDICION = True Then
            Check1(5).Value = Checked
        End If
        If .getPER_CIERRE = True Then
            Check1(6).Value = Checked
        End If
        If .getPER_PEDIDOS_REACTIVOS = True Then
            Check1(7).Value = Checked
        End If
        If .getPER_EMPLEADOS = True Then
            Check1(8).Value = Checked
        End If
        If .getPER_CONTABILIDAD = True Then
            Check1(9).Value = Checked
        End If
        If .getPER_ENVIO_PEDIDOS_PROVEEDOR = True Then
            Check1(10).Value = Checked
        End If
        If .getPER_DOCUMENTACION_CALIDAD = True Then
            Check1(11).Value = Checked
        End If
        If .getPER_PROYECTOS = True Then
            Check1(12).Value = Checked
        End If
        If .getPER_NC = True Then
            Check1(13).Value = Checked
        End If
        If .getPER_PNT = True Then
            Check1(15).Value = Checked
        End If
        If .getPER_ADMIN_PNT = True Then
            Check1(16).Value = Checked
        End If
        If .getPER_DATOS_ESPECIALES = True Then
            Check1(17).Value = Checked
        End If
        If .getPER_MATRIZ_CUALIF = True Then
            Check1(18).Value = Checked
        End If
        If .getPER_VIDEOS = True Then
            Check1(19).Value = Checked
        End If
        If .getPER_IMPRESION_PNT = True Then
            Check1(20).Value = Checked
        End If
        If .getPER_INCURRIDOS = True Then
            Check1(21).Value = Checked
        End If
        If .getPER_REVISION = True Then
            Check1(22).Value = Checked
        End If
        If .getPER_OFERTAS = True Then
            Check1(23).Value = Checked
        End If
        If .getPER_PLAZO_ENTREGA_CAMBIO = True Then
            Check1(27).Value = Checked
        End If
        If .getPER_PLAZO_ENTREGA_LISTADO = True Then
            Check1(25).Value = Checked
        End If
        If .getPER_INDICADORES_CLIENTE = True Then
            Check1(28).Value = Checked
        End If
        If .getPER_FAMILIAS_CA = True Then
            Check1(24).Value = Checked
        End If
        'M1144-I
        If .getPER_TRAMITACION_CONTRATA = True Then
            Check1(26).Value = Checked
        End If
        If .getPER_RFI = True Then
            Check1(29).Value = Checked
        End If
        If .getPER_PFA = True Then
            Check1(30).Value = Checked
        End If
        'M1144-F
        If .getPER_SCG = True Then
            Check1(31).Value = Checked
        End If
        If .getPER_TESORERIA_MENU = True Then
            Check1(32).Value = Checked
        End If
        If .getPER_TESORERIA_FP = True Then
            Check1(33).Value = Checked
        End If
        'M1377-I
        If .getPER_NORMAS_NO_CONTROLADAS = True Then
            Check1(34).Value = Checked
        End If
        If .getPER_DES_PRODUCTO = True Then
            Check1(35).Value = Checked
        End If
        'M1377-F
        If .getPER_MOD_CLIENTE = True Then
            Check1(36).Value = Checked
        End If
        If .getPER_INCIDENCIAS = True Then
            Check1(37).Value = Checked
        End If
        If .getPER_PRODUCTIVIDAD = True Then
            Check1(38).Value = Checked
        End If
        
        'JONATHAN.2009.10.27
        For intCont = 1 To TOTAL_DEPARTAMENTOS
            chkDepartamento(intCont).Value = .getRESPONSABLE_DEPARTAMENTOS(intCont)
            mvarstrResposabilidadesDepartamento(intCont) = .getRESPONSABLE_DEPARTAMENTOS(intCont)
        Next intCont
        'FIN JONATHAN
        
        'JONATHAN.2009.10.17
        
        For Each objDept In .getDEPARTAMENTOS_USUARIO.Iterator
            chkDepartamentoUsuario(objDept.getID).Value = vbChecked
        Next objDept
        Set objDept = Nothing
        Set mvarobjDepartamentosUsuario = .getDEPARTAMENTOS_USUARIO
        'FIN JONATHAN
        
        
        
        datos(4) = Replace(.getFIRMA, "/", "\")
        datos(5) = Replace(.getIMAGEN, "/", "\")
        'E0145-I
        datos(6) = .getEMAIL
        datos(7) = Replace(.getFNMT_RUTA, "/", "\")
        datos(8) = .getFNMT_PASS
        'E0145-F
        lbltitulo.Caption = "Modificacion del usuario : " & .getUSUARIO
     End With
    End If
End Sub
Private Sub cmdEXplorar_Click()
    cd.DialogTitle = "Abrir fichero de imagen"
    cd.InitDir = ReadINI(App.Path + "\config.ini", "Documentos", "Firmas")
    cd.ShowOpen
    If cd.FileName <> "" Then
        datos(4).Text = cd.FileName  ' cd.FileTitle
    End If
End Sub
Private Sub datos_Change(Index As Integer)
    On Error Resume Next
    If Index = 4 And datos(4) <> "" Then
        If Dir(datos(4)) <> "" Then
            Set img.Picture = LoadPicture(datos(4))
        End If
    End If
    If Index = 5 And datos(5) <> "" Then
        If Dir(datos(5)) <> "" Then
            Set Image3.Picture = LoadPicture(datos(5))
        End If
    End If

End Sub


Private Sub permisos()
    If UCase(USUARIO.getUSUARIO) = "JULIO" Then
        Check1(16).visible = True
        Check1(22).Enabled = True
    Else
        Check1(16).visible = False
        Check1(22).Enabled = False
    End If
End Sub
