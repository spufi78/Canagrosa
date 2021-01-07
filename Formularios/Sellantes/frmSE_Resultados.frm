VERSION 5.00
Object = "{0E59F1D2-1FBE-11D0-8FF2-00A0D10038BC}#1.0#0"; "msscript.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{CEA3B7F6-2847-4E5E-A551-DB7A62489D44}#46.0#0"; "miCombo.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#13.2#0"; "Codejock.Controls.v13.2.1.ocx"
Begin VB.Form frmSE_Resultados 
   BackColor       =   &H00C0C0C0&
   Caption         =   "Resultados de Sellante"
   ClientHeight    =   10005
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   13710
   Icon            =   "frmSE_Resultados.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   10005
   ScaleWidth      =   13710
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdVerSellante 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Ver Sellante"
      Height          =   840
      Left            =   45
      Style           =   1  'Graphical
      TabIndex        =   92
      Top             =   9090
      Width           =   1140
   End
   Begin VB.CommandButton cmdImagen 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Imagenes"
      Height          =   840
      Left            =   10260
      Style           =   1  'Graphical
      TabIndex        =   88
      Top             =   9090
      Width           =   1140
   End
   Begin VB.CommandButton cmdObservador 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Observador"
      Height          =   840
      Left            =   11407
      Style           =   1  'Graphical
      TabIndex        =   87
      Tag             =   "Añade campo o modifica el campo existente con el mismo nombre"
      Top             =   9090
      Width           =   1140
   End
   Begin VB.CheckBox chkModificar 
      Caption         =   "Permiso Modificar Cerrada"
      Height          =   195
      Left            =   5715
      TabIndex        =   86
      Top             =   9450
      Visible         =   0   'False
      Width           =   1230
   End
   Begin VB.TextBox txtID_SELLANTE 
      Height          =   285
      Left            =   3915
      TabIndex        =   82
      Text            =   "0"
      Top             =   9225
      Visible         =   0   'False
      Width           =   915
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   600
      Left            =   9315
      TabIndex        =   78
      Top             =   9090
      Visible         =   0   'False
      Width           =   780
   End
   Begin MSComctlLib.ListView auxdatos 
      Height          =   3555
      Left            =   8055
      TabIndex        =   74
      Top             =   4590
      Visible         =   0   'False
      Width           =   4650
      _ExtentX        =   8202
      _ExtentY        =   6271
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
   Begin VB.CheckBox chkDuplicada 
      Caption         =   "Duplicada"
      Height          =   195
      Left            =   7245
      TabIndex        =   73
      Top             =   9360
      Visible         =   0   'False
      Width           =   1050
   End
   Begin VB.Frame Frame5 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   "Datos del Ensayo"
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
      Height          =   1380
      Left            =   7965
      TabIndex        =   69
      Top             =   1830
      Width           =   5685
      Begin VB.CommandButton cmdModificarEnsayo 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Modificar"
         Height          =   930
         Left            =   4455
         Picture         =   "frmSE_Resultados.frx":08CA
         Style           =   1  'Graphical
         TabIndex        =   18
         Top             =   270
         Width           =   1140
      End
      Begin VB.TextBox txtdatos 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   9
         Left            =   3015
         TabIndex        =   17
         Top             =   945
         Width           =   1275
      End
      Begin VB.TextBox txtdatos 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   14
         Left            =   3015
         TabIndex        =   16
         Top             =   630
         Width           =   1275
      End
      Begin VB.TextBox txtdatos 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   11
         Left            =   3015
         TabIndex        =   15
         Top             =   300
         Width           =   1275
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
         Index           =   13
         Left            =   135
         TabIndex        =   72
         Top             =   945
         Width           =   870
      End
      Begin VB.Label Label1 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Humedad Relativa (40-60% Hr):"
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
         Index           =   17
         Left            =   135
         TabIndex        =   71
         Top             =   675
         Width           =   2895
      End
      Begin VB.Label Label1 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Temperatura (21-25º):"
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
         Index           =   15
         Left            =   150
         TabIndex        =   70
         Top             =   345
         Width           =   2115
      End
   End
   Begin Geslab.ControlPanelXP ControlPanelXP2 
      Height          =   3840
      Left            =   45
      TabIndex        =   35
      Top             =   3870
      Width           =   6810
      _ExtentX        =   12012
      _ExtentY        =   6773
      Caption         =   "Equipos Utilizados en la Muestra"
      BackColor       =   12632256
      TextColor       =   0
      HeaderColor     =   8421504
      PanelOpen       =   0   'False
      Object.Height          =   3840
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
         Height          =   3390
         Left            =   90
         TabIndex        =   36
         Top             =   405
         Width           =   6585
         Begin VB.CommandButton cmdVerificacion 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Verificación"
            Enabled         =   0   'False
            Height          =   765
            Left            =   5670
            Style           =   1  'Graphical
            TabIndex        =   77
            Tag             =   "Añade campo o modifica el campo existente con el mismo nombre"
            Top             =   1920
            Width           =   915
         End
         Begin VB.CommandButton cmdAnadirEquipo 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Añadir"
            Height          =   765
            Left            =   5670
            Style           =   1  'Graphical
            TabIndex        =   38
            Tag             =   "Añade campo o modifica el campo existente con el mismo nombre"
            Top             =   1110
            Width           =   915
         End
         Begin VB.CommandButton cmdEliminarEquipo 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Eliminar"
            Height          =   810
            Left            =   5670
            Style           =   1  'Graphical
            TabIndex        =   37
            Tag             =   "Elimina el campo seleccionado"
            Top             =   270
            Width           =   915
         End
         Begin MSComctlLib.ListView listaEquipos 
            Height          =   2325
            Left            =   0
            TabIndex        =   39
            Top             =   270
            Width           =   5580
            _ExtentX        =   9843
            _ExtentY        =   4101
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
            TabIndex        =   40
            Top             =   2700
            Width           =   5595
            _ExtentX        =   9869
            _ExtentY        =   582
         End
         Begin VB.Label Label2 
            BackColor       =   &H00C0C0C0&
            Caption         =   "Marque los equipos que deben salir en el informe"
            ForeColor       =   &H000000FF&
            Height          =   195
            Left            =   0
            TabIndex        =   41
            Top             =   45
            Width           =   4335
         End
      End
   End
   Begin Geslab.ControlPanelXP cpReactivos 
      Height          =   3975
      Left            =   6885
      TabIndex        =   42
      Top             =   3870
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
         TabIndex        =   43
         Top             =   450
         Width           =   6630
         Begin VB.CommandButton cmdAnadirReactivo 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Añadir"
            Height          =   750
            Left            =   5670
            Style           =   1  'Graphical
            TabIndex        =   45
            Tag             =   "Añade campo o modifica el campo existente con el mismo nombre"
            Top             =   1395
            Width           =   915
         End
         Begin VB.CommandButton cmdEliminarReactivo 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Eliminar"
            Height          =   795
            Left            =   5670
            Style           =   1  'Graphical
            TabIndex        =   44
            Tag             =   "Elimina el campo seleccionado"
            Top             =   450
            Width           =   915
         End
         Begin MSComctlLib.ListView listaReactivos 
            Height          =   2460
            Left            =   45
            TabIndex        =   46
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
            TabIndex        =   47
            Top             =   2700
            Width           =   5280
            _ExtentX        =   9313
            _ExtentY        =   582
         End
         Begin pryCombo.miCombo cmbReactivosInternos 
            Height          =   330
            Left            =   765
            TabIndex        =   48
            Top             =   3060
            Width           =   5280
            _ExtentX        =   9313
            _ExtentY        =   582
         End
         Begin VB.Label lblCampos 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C0C0&
            Caption         =   "Externo"
            Height          =   195
            Index           =   5
            Left            =   90
            TabIndex        =   50
            Top             =   2745
            Width           =   540
         End
         Begin VB.Label lblCampos 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C0C0&
            Caption         =   "Interno"
            Height          =   195
            Index           =   8
            Left            =   90
            TabIndex        =   49
            Top             =   3105
            Width           =   495
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
      Height          =   4785
      Left            =   45
      TabIndex        =   51
      Top             =   4275
      Width           =   13605
      Begin VB.TextBox txtdatos 
         Appearance      =   0  'Flat
         Height          =   465
         Index           =   10
         Left            =   45
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   80
         Top             =   3240
         Width           =   7080
      End
      Begin VB.Frame Frame2 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         ForeColor       =   &H80000008&
         Height          =   1080
         Left            =   45
         TabIndex        =   61
         Top             =   3645
         Width           =   7095
         Begin VB.TextBox txtvalor 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   1260
            TabIndex        =   2
            Top             =   630
            Width           =   2700
         End
         Begin VB.TextBox txtdato 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   1260
            Locked          =   -1  'True
            TabIndex        =   65
            Top             =   240
            Width           =   2700
         End
         Begin VB.CommandButton cmdAceptar 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Aceptar"
            Height          =   840
            Left            =   5895
            Style           =   1  'Graphical
            TabIndex        =   64
            Top             =   180
            Width           =   1140
         End
         Begin VB.OptionButton chkConforme 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            Caption         =   "Conforme"
            ForeColor       =   &H80000008&
            Height          =   195
            Index           =   0
            Left            =   4365
            TabIndex        =   63
            Top             =   360
            Width           =   1050
         End
         Begin VB.OptionButton chkConforme 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            Caption         =   "NO Conforme"
            ForeColor       =   &H80000008&
            Height          =   195
            Index           =   1
            Left            =   4365
            TabIndex        =   62
            Top             =   720
            Width           =   1275
         End
         Begin VB.Label Label1 
            BackColor       =   &H00C0C0C0&
            Caption         =   "Resultado"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   255
            Index           =   3
            Left            =   150
            TabIndex        =   67
            Top             =   690
            Width           =   975
         End
         Begin VB.Label Label1 
            BackColor       =   &H00C0C0C0&
            Caption         =   "Ensayo"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   255
            Index           =   4
            Left            =   135
            TabIndex        =   66
            Top             =   300
            Width           =   975
         End
      End
      Begin VB.Frame Frame4 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         ForeColor       =   &H80000008&
         Height          =   720
         Left            =   7200
         TabIndex        =   52
         Top             =   3825
         Width           =   6360
         Begin VB.TextBox txtdato2 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            Height          =   330
            Left            =   855
            Locked          =   -1  'True
            TabIndex        =   55
            Top             =   225
            Width           =   2355
         End
         Begin VB.TextBox txtvalor2 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            Height          =   330
            Left            =   3825
            TabIndex        =   53
            Top             =   225
            Width           =   1635
         End
         Begin VB.CommandButton cmdcalcular 
            BackColor       =   &H00E0E0E0&
            Enabled         =   0   'False
            Height          =   555
            Left            =   5670
            Picture         =   "frmSE_Resultados.frx":1194
            Style           =   1  'Graphical
            TabIndex        =   54
            Top             =   135
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
            Index           =   12
            Left            =   90
            TabIndex        =   57
            Top             =   315
            Width           =   855
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
            Index           =   5
            Left            =   3285
            TabIndex        =   56
            Top             =   270
            Width           =   555
         End
      End
      Begin MSComctlLib.ListView datos 
         Height          =   3240
         Left            =   7200
         TabIndex        =   58
         Top             =   495
         Width           =   6300
         _ExtentX        =   11113
         _ExtentY        =   5715
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
      Begin MSComctlLib.ListView lista 
         Height          =   2100
         Left            =   45
         TabIndex        =   1
         Top             =   495
         Width           =   7050
         _ExtentX        =   12435
         _ExtentY        =   3704
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         AllowReorder    =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         Icons           =   "ImageList1"
         SmallIcons      =   "ImageList1"
         ColHdrIcons     =   "ImageList1"
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
      Begin XtremeSuiteControls.PushButton cmdEnsayoAdd 
         Height          =   345
         Left            =   4770
         TabIndex        =   83
         Top             =   2610
         Width           =   1140
         _Version        =   851970
         _ExtentX        =   2011
         _ExtentY        =   609
         _StockProps     =   79
         Caption         =   "Añadir"
         Appearance      =   5
         Picture         =   "frmSE_Resultados.frx":149E
      End
      Begin XtremeSuiteControls.PushButton cmdEnsayoEliminar 
         Height          =   345
         Left            =   5940
         TabIndex        =   84
         Top             =   2610
         Width           =   1140
         _Version        =   851970
         _ExtentX        =   2011
         _ExtentY        =   609
         _StockProps     =   79
         Caption         =   "Eliminar"
         Appearance      =   5
         Picture         =   "frmSE_Resultados.frx":7D00
      End
      Begin pryCombo.miCombo cmbEnsayos 
         Height          =   375
         Left            =   45
         TabIndex        =   85
         Top             =   2610
         Width           =   4710
         _ExtentX        =   8308
         _ExtentY        =   661
      End
      Begin VB.Label Label1 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Criterio Aceptación"
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
         Index           =   14
         Left            =   45
         TabIndex        =   81
         Top             =   3015
         Width           =   2670
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         Caption         =   "Resultados"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   300
         Index           =   1
         Left            =   45
         TabIndex        =   68
         Top             =   180
         Width           =   7035
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
         Left            =   10800
         TabIndex        =   60
         Top             =   180
         Width           =   2715
      End
      Begin VB.Label Label5 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Campos de tipo determinacion"
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
         Height          =   285
         Index           =   0
         Left            =   7200
         TabIndex        =   59
         Top             =   180
         Width           =   6285
      End
   End
   Begin VB.Frame Frame3 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   "Datos de la recepción"
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
      Height          =   2010
      Left            =   60
      TabIndex        =   27
      Top             =   1830
      Width           =   7890
      Begin VB.TextBox txtdatos 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   12
         Left            =   4635
         TabIndex        =   7
         Top             =   210
         Width           =   1920
      End
      Begin VB.CheckBox chkfMezcla 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Sin Especificar"
         Height          =   255
         Left            =   2700
         TabIndex        =   79
         Top             =   1215
         Width           =   1485
      End
      Begin VB.CheckBox chkFLimite 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Sin Especificar"
         Height          =   255
         Left            =   2700
         TabIndex        =   76
         Top             =   1590
         Width           =   1485
      End
      Begin VB.CommandButton cmdModificar 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Modificar"
         Height          =   930
         Left            =   6660
         Style           =   1  'Graphical
         TabIndex        =   14
         Top             =   300
         Width           =   1140
      End
      Begin VB.TextBox txtdatos 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   4
         Left            =   1365
         TabIndex        =   6
         Top             =   210
         Width           =   1920
      End
      Begin VB.TextBox txtdatos 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   5
         Left            =   1365
         TabIndex        =   8
         Top             =   525
         Width           =   5195
      End
      Begin VB.TextBox txtdatos 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   7
         Left            =   1365
         TabIndex        =   9
         Top             =   840
         Width           =   1920
      End
      Begin VB.TextBox txtdatos 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Index           =   6
         Left            =   4635
         TabIndex        =   10
         Top             =   840
         Width           =   1920
      End
      Begin VB.TextBox txtdatos 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Index           =   8
         Left            =   5625
         TabIndex        =   12
         Top             =   1200
         Width           =   945
      End
      Begin MSComCtl2.DTPicker fecha 
         Height          =   330
         Left            =   1380
         TabIndex        =   11
         Top             =   1170
         Width           =   1305
         _ExtentX        =   2302
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
         Format          =   51576833
         CurrentDate     =   38000
         MinDate         =   2
      End
      Begin MSComCtl2.DTPicker fecha_limite 
         Height          =   330
         Left            =   1380
         TabIndex        =   13
         Top             =   1575
         Width           =   1305
         _ExtentX        =   2302
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
         Format          =   51576833
         CurrentDate     =   38000
         MinDate         =   2
      End
      Begin VB.Label Label1 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Ratio"
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
         Index           =   16
         Left            =   3465
         TabIndex        =   89
         Top             =   225
         Width           =   1140
      End
      Begin VB.Label lblCampos 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "F. Límite"
         Height          =   195
         Index           =   0
         Left            =   105
         TabIndex        =   75
         Top             =   1620
         Width           =   645
      End
      Begin VB.Label lblCampos 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Fecha Mezcla"
         Height          =   195
         Index           =   6
         Left            =   105
         TabIndex        =   33
         Top             =   1230
         Width           =   1035
      End
      Begin VB.Label Label1 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Nº Mezcla"
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
         Index           =   11
         Left            =   105
         TabIndex        =   32
         Top             =   255
         Width           =   1140
      End
      Begin VB.Label Label1 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Nº Lote y Kit"
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
         Index           =   10
         Left            =   105
         TabIndex        =   31
         Top             =   570
         Width           =   1215
      End
      Begin VB.Label Label1 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Temperatura"
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
         Index           =   9
         Left            =   105
         TabIndex        =   30
         Top             =   885
         Width           =   1440
      End
      Begin VB.Label Label1 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Higrometría"
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
         Index           =   8
         Left            =   3450
         TabIndex        =   29
         Top             =   870
         Width           =   1230
      End
      Begin VB.Label Label1 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Hora "
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
         Height          =   225
         Index           =   6
         Left            =   5115
         TabIndex        =   28
         Top             =   1260
         Width           =   570
      End
   End
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   "Datos del Sellante"
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
      Height          =   1425
      Left            =   45
      TabIndex        =   21
      Top             =   375
      Width           =   13605
      Begin VB.TextBox txtdatos 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         Height          =   285
         Index           =   3
         Left            =   6750
         Locked          =   -1  'True
         TabIndex        =   26
         Top             =   900
         Width           =   5280
      End
      Begin VB.TextBox txtdatos 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         Height          =   285
         Index           =   2
         Left            =   1365
         Locked          =   -1  'True
         TabIndex        =   5
         Top             =   945
         Width           =   4380
      End
      Begin VB.TextBox txtdatos 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         Height          =   285
         Index           =   1
         Left            =   6750
         Locked          =   -1  'True
         TabIndex        =   4
         Top             =   585
         Width           =   5280
      End
      Begin VB.TextBox txtdatos 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         Height          =   285
         Index           =   0
         Left            =   1365
         Locked          =   -1  'True
         TabIndex        =   3
         Top             =   585
         Width           =   4380
      End
      Begin MSDataListLib.DataCombo cmbproducto 
         Height          =   315
         Left            =   1350
         TabIndex        =   0
         Top             =   225
         Width           =   10695
         _ExtentX        =   18865
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
      Begin VB.CommandButton cmdModificarSellante 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Modificar"
         Height          =   930
         Left            =   12330
         Picture         =   "frmSE_Resultados.frx":E562
         Style           =   1  'Graphical
         TabIndex        =   91
         Top             =   270
         Width           =   1140
      End
      Begin VB.Label Label1 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Producto"
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
         Index           =   18
         Left            =   135
         TabIndex        =   90
         Top             =   315
         Width           =   810
      End
      Begin VB.Label Label1 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Producto"
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
         Index           =   7
         Left            =   5895
         TabIndex        =   25
         Top             =   945
         Width           =   915
      End
      Begin VB.Label Label1 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Preparación"
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
         Left            =   135
         TabIndex        =   24
         Top             =   945
         Width           =   1440
      End
      Begin VB.Label Label1 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Proceso"
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
         Left            =   5895
         TabIndex        =   23
         Top             =   630
         Width           =   810
      End
      Begin VB.Label Label1 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Ensayo"
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
         Index           =   0
         Left            =   135
         TabIndex        =   22
         Top             =   630
         Width           =   810
      End
   End
   Begin VB.CommandButton cmdSalir 
      BackColor       =   &H00E0E0E0&
      Caption         =   "ESC-Salir"
      Height          =   840
      Left            =   12555
      Style           =   1  'Graphical
      TabIndex        =   19
      Top             =   9090
      Width           =   1140
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   8415
      Top             =   9270
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   1
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSE_Resultados.frx":EE2C
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSScriptControlCtl.ScriptControl sc 
      Left            =   2790
      Top             =   9135
      _ExtentX        =   1005
      _ExtentY        =   1005
   End
   Begin VB.Label lblCerrada 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
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
      Height          =   345
      Left            =   12195
      TabIndex        =   34
      Top             =   0
      Width           =   1500
   End
   Begin VB.Label lbltitulo 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      Caption         =   "Resultados de Sellante"
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
      TabIndex        =   20
      Top             =   0
      Width           =   13545
   End
End
Attribute VB_Name = "frmSE_Resultados"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmbproducto_Change()
    Dim oSellante As New clsSellantes
    With oSellante
        .Carga cmbProducto.BoundText
        txtDatos(0) = .getENSAYO
        txtDatos(1) = .getPROCESO
        txtDatos(2) = .getPREPARACION
        txtDatos(3) = .getPRODUCTO
    End With
    txtID_SELLANTE = cmbProducto.BoundText
End Sub

Private Sub cmdImagen_Click()
    With frmCE_Imagenes
        .PK = gmuestra
        .Show 1
    End With
End Sub
Private Sub chkfmezcla_Click()
    If chkfMezcla.Value = Checked Then
        fecha.Enabled = False
    Else
        fecha.Enabled = True
    End If
End Sub
Private Sub cmdEnsayoAdd_Click()
    If cmbEnsayos.getTEXTO <> "" Then
        Dim oSe As New clsSellantes_ensayos
        Dim oSR As New clsSellantes_resultados
        
        oSe.Carga txtID_SELLANTE, cmbEnsayos.getPK_SALIDA
        With oSR
            .setMUESTRA_ID = gmuestra
            .setSELLANTE_ID = txtID_SELLANTE
            .setORDEN = cmbEnsayos.getPK_SALIDA
            .setTIPO_DETERMINACION_ID = oSe.getTIPO_DETERMINACION_ID
            If oSe.getTIPO_DETERMINACION_ID <> 0 Then
                Dim oTD As New clsTipos_determinacion
                oTD.CargarTipoDeterminacion oSe.getTIPO_DETERMINACION_ID
                .setFORMULA_ID = oTD.getFORMULA_ID
            Else
                .setFORMULA_ID = 0
            End If
            .setRESULTADO = ""
            .setCONFORME = 0
            .Insertar
        End With
        cargar_resultados
    End If
End Sub

Private Sub cmdEnsayoEliminar_Click()
    If lista.ListItems.Count > 0 Then
        If MsgBox("¿Esta seguro de eliminar : " & lista.ListItems(lista.selectedItem.Index).SubItems(1), vbQuestion + vbYesNo) = vbYes Then
            Dim oSR As New clsSellantes_resultados
            oSR.Eliminar gmuestra, lista.ListItems(lista.selectedItem.Index).Text
            Set oSR = Nothing
            cargar_resultados
        End If
    End If
End Sub

Private Sub cmdModificarSellante_Click()
   On Error GoTo cmdModificarSellante_Click_Error

    If MsgBox("¿Modificar el Sellante del ensayo?", vbYesNo + vbQuestion, App.Title) = vbYes Then
        Dim oSellante_recepcion As New clsSellantes_recepcion
        oSellante_recepcion.ModificarSellante gmuestra, CLng(txtID_SELLANTE)
        Set oSellante_recepcion = Nothing
        Dim oSellante_resultados As New clsSellantes_resultados
        oSellante_resultados.ModificarSellante gmuestra, CLng(txtID_SELLANTE)
        Set oSellante_resultados = Nothing
        MsgBox "Modificación realizada correctamente.", vbInformation + vbOKOnly, App.Title
    End If

   On Error GoTo 0
   Exit Sub

cmdModificarSellante_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure cmdModificarSellante_Click of Formulario frmSE_Resultados"
End Sub

'MANTIS-807-I
Private Sub cmdObservador_Click()

    Dim objfrm As New frmObservadorEnsayo

    objfrm.FORMULARIO_ORIGEN = 2 'Sellantes asociado al número 2
    objfrm.ES_CONTROL_EFICACIA = False
    objfrm.MUESTRA_ID = gmuestra ' Id de la muestra
    objfrm.DETERMINACION_ENSAYO_ID = 0
'M0961-I
    objfrm.SELLANTE_ID = txtID_SELLANTE
    objfrm.ENSAYO = lista.ListItems(lista.selectedItem.Index)
'M0961-F
    
    If (UCase(lblCerrada) <> "CERRADA") Then
        objfrm.MUESTRA_CERRADA = False
    Else
        objfrm.MUESTRA_CERRADA = True
    End If

    objfrm.Show vbModal
    
    Set objfrm = Nothing

End Sub
'MANTIS-807-F'

Private Sub cmdVerificacion_Click()
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
            'objfrm.IdPeriodoInicial = ENUM_EQUIPOS_PERIODICIDAD.ENUM_EQUIPOS_PERIODICIDAD_ANTES_ENSAYO
            objfrm.IdPeriodoInicial = ENUM_EQUIPOS_PERIODICIDAD.ENUM_EQUIPOS_PERIODICIDAD_ANTES_CADA_ENSAYO
            'MANTIS-810-F
            objfrm.IdTipoVerificacionIncial = 1
            
            'MANTIS-810-I
            'objfrm.copiarUltimaVerificacionPeriodo = ENUM_EQUIPOS_PERIODICIDAD.ENUM_EQUIPOS_PERIODICIDAD_ANTES_ENSAYO
            objfrm.copiarUltimaVerificacionPeriodo = ENUM_EQUIPOS_PERIODICIDAD.ENUM_EQUIPOS_PERIODICIDAD_ANTES_CADA_ENSAYO
            'MANTIS-810-F
              
            objfrm.Show vbModal
            If objfrm.ID_VERIFICACION <> 0 Then
                listaEquipos.ListItems(listaEquipos.selectedItem.Index).SubItems(3) = objfrm.ID_VERIFICACION
            End If
            grabar_equipos
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
End Sub
Private Sub chkFLimite_Click()
    If chkFLimite.Value = Checked Then
        fecha_limite.Enabled = False
    Else
        fecha_limite.Enabled = True
    End If
End Sub

Private Sub chkConforme_Click(Index As Integer)
    If Trim(txtValor) = "CONFORME" Or Trim(txtValor) = "NO CONFORME" Or Trim(txtValor) = "" Then
        If Index = 0 Then
            txtValor = "CONFORME"
        Else
            txtValor = "NO CONFORME"
        End If
    End If
End Sub
Private Sub cmdAceptar_Click()
    If lista.ListItems.Count > 0 Then
        insertar_resultado (False)
    End If
End Sub
Private Sub insertar_resultado(es_determinacion As Boolean)
    If txtValor = "" Then
        If chkConforme(0).Value = False And chkConforme(1).Value = False Then
            MsgBox "Introduzca resultado númerico, o Conforme - No conforme.", vbExclamation, App.Title
            Exit Sub
        End If
    Else
        ' Validar rangos
        If IsNumeric(txtValor) Then
            lista.ListItems(lista.selectedItem.Index).SubItems(6) = "1"
            If Not es_determinacion Then
                ' minimo
                If Trim(lista.ListItems(lista.selectedItem.Index).SubItems(2)) <> "" Then
                    If CSng(txtValor) < CSng(lista.ListItems(lista.selectedItem.Index).SubItems(2)) Then
                        MsgBox "ATENCION: El valor introducido es MENOR que el mínimo establecido.", vbExclamation, App.Title
                        lista.ListItems(lista.selectedItem.Index).SubItems(6) = "0"
                    End If
                End If
                ' maximo
                If Trim(lista.ListItems(lista.selectedItem.Index).SubItems(3)) <> "" Then
                    If CSng(txtValor) > CSng(lista.ListItems(lista.selectedItem.Index).SubItems(3)) Then
                        MsgBox "ATENCION: El valor introducido es MAYOR que el máximo establecido.", vbExclamation, App.Title
                        lista.ListItems(lista.selectedItem.Index).SubItems(6) = "0"
                    End If
                End If
            End If
        End If
    End If
    ' Grbar Resultados
    lista.ListItems(lista.selectedItem.Index).SubItems(4) = txtValor
    If chkConforme(0).Value = True Then
        lista.ListItems(lista.selectedItem.Index).SubItems(6) = "1"
    End If
    If chkConforme(1).Value = True Then
        lista.ListItems(lista.selectedItem.Index).SubItems(6) = "0"
    End If
    ' Si no es conforme se pone en rojo
    colorear_linea (lista.selectedItem.Index)
    ' Grabar en la bd
    If UCase(lblCerrada) <> "CERRADA" Then
        Dim oSe_resultado As New clsSellantes_resultados
        With oSe_resultado
            .setRESULTADO = lista.ListItems(lista.selectedItem.Index).SubItems(4)
            .setCONFORME = lista.ListItems(lista.selectedItem.Index).SubItems(6)
            .informar_resultado gmuestra, lista.ListItems(lista.selectedItem.Index)
        End With
    End If
    ' Pasar al siguiente campo
    If Not es_determinacion Then
        If lista.ListItems.Count > lista.selectedItem.Index Then
            Set lista.selectedItem = lista.ListItems(lista.selectedItem.Index + 1)
            lista.selectedItem.EnsureVisible
            lista_Click
        End If
    End If
End Sub

Private Sub cmdAnadirEquipo_Click()
    If cmbEquipos.getPK_SALIDA <> 0 Then
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
            .SubItems(3) = "0"
        End With
        listaEquipos.ListItems(listaEquipos.ListItems.Count).EnsureVisible
        cmbEquipos.limpiar
        grabar_equipos
    End If

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
            .SubItems(1) = oTRPR.getNOMBRE
            .SubItems(2) = Format(oRPR.getFECHA_CADUCIDAD, "dd-mm-yyyy")
            .SubItems(3) = "I"
        End With
        listaReactivos.ListItems(listaReactivos.ListItems.Count).EnsureVisible
    End If
    ' Limpiar Combos
    cmbReactivos.limpiar
    cmbReactivosInternos.limpiar
    grabar_reactivos
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
    Dim pos As Integer
    Dim ofor As New clsFormulas
    Dim encontrado As Boolean
    Dim j As Integer
'    prefijo = ""
'    Dim oTD As New clsTipos_determinacion
'    oTD.CargarTipoDeterminacion (lista.ListItems(lista.SelectedItem.Index).SubItems(7))
'    ofor.CARGAR (oTD.getFORMULA_ID)
    ofor.CARGAR (lista.ListItems(lista.selectedItem.Index).SubItems(8))
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
                pos = InStr(i + 2, cadena, "_")
                campo = Mid(cadena, i + 2, (pos) - (i + 2))
                j = datos.selectedItem.Index
                encontrado = False
                Do
                 If CInt(datos.ListItems(j).SubItems(3)) = CInt(campo) Then
                     Formula = Formula & Replace(datos.ListItems(j).SubItems(1), ",", ".")
                     encontrado = True
                 End If
                 j = j - 1
                Loop Until j = 0 Or encontrado = True
                i = pos
            End If
        Next
    End If
    If Formula <> "" Then
        datos.ListItems(datos.selectedItem.Index).SubItems(1) = formatear(sc.Eval(Formula), datos.ListItems(datos.selectedItem.Index).SubItems(4), datos.ListItems(datos.selectedItem.Index).SubItems(5))
    End If
    grabar_auxdatos
    visualizar_duplicados
    pasar_siguiente_campo
    Exit Sub
fallo:
    MsgBox "Error en la formula. " & Err.Description, vbCritical, "Error"

End Sub

Private Sub cmdEliminarEquipo_Click()
    If listaEquipos.ListItems.Count > 0 Then
        listaEquipos.ListItems.Remove listaEquipos.selectedItem.Index
        grabar_equipos
    End If
End Sub

Private Sub cmdEliminarReactivo_Click()
    If listaReactivos.ListItems.Count > 0 Then
        listaReactivos.ListItems.Remove listaReactivos.selectedItem.Index
        cmbReactivos.limpiar
        cmbReactivosInternos.limpiar
    End If
    grabar_reactivos
End Sub

Private Sub cmdModificar_Click()
   On Error GoTo cmdModificar_Click_Error

    If MsgBox("¿Modificar los datos del registro?", vbYesNo + vbQuestion, App.Title) = vbYes Then
        Dim oSellante_recepcion As New clsSellantes_recepcion
        With oSellante_recepcion
            If chkfMezcla.Value = Checked Then
                .setFECHA = "0000-00-00"
            Else
                .setFECHA = Format(fecha, "yyyy-mm-dd")
            End If
            If chkFLimite.Value = Checked Then
                .setFECHA_LIMITE = "0000-00-00"
            Else
                .setFECHA_LIMITE = Format(fecha_limite, "yyyy-mm-dd")
            End If
            .setN_MEZCLA = txtDatos(4)
            .setR_MEZCLA = txtDatos(12)
            .setLOTE = txtDatos(5)
            .setHIGROMETRIA = txtDatos(6)
            .setTEMPERATURA = txtDatos(7)
            .setHORA = txtDatos(8)
            .Modificar (gmuestra)
        End With
        MsgBox "Modificación realizada correctamente.", vbInformation + vbOKOnly, App.Title
    End If

   On Error GoTo 0
   Exit Sub

cmdModificar_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure cmdModificar_Click of Formulario frmSE_Resultados"
End Sub

Private Sub cmdModificarEnsayo_Click()

   On Error GoTo cmdModificarEnsayo_Click_Error

    If MsgBox("¿Modificar los datos del ensayo?", vbYesNo + vbQuestion, App.Title) = vbYes Then
        Dim oSellante_recepcion As New clsSellantes_recepcion
        With oSellante_recepcion
            .setENSAYO_TEMPERATURA = txtDatos(11)
            .setENSAYO_HUMEDAD = txtDatos(14)
            .setENSAYO_ESPESOR = txtDatos(9)
            .ModificarEnsayo (gmuestra)
        End With
        MsgBox "Modificación realizada correctamente.", vbInformation + vbOKOnly, App.Title
    End If

   On Error GoTo 0
   Exit Sub

cmdModificarEnsayo_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure cmdModificarEnsayo_Click of Formulario frmSE_Resultados"

End Sub

Private Sub cmdSalir_Click()
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
        If MsgBox("Existen reactivos CADUCADOS. ¿ESTA SEGURO DE ALMACENAR LOS DATOS DE LA MUESTRA?", vbExclamation + vbYesNo, App.Title) = vbNo Then
            Exit Sub
        End If
    End If

    grabar_equipos
    Dim oMuestra As New clsMuestra
    oMuestra.comprobar_cierre (gmuestra)
    Unload Me
End Sub


Private Sub cmdVerSellante_Click()
    gSE_Sellante = txtID_SELLANTE
    frmSE_Detalle.Show 1
End Sub

Private Sub Command1_Click()
    execute_bd "UPDATE SELLANTES_DETERMINACIONES " & _
               "   SET VALOR_1 = '" & Replace(auxdatos.ListItems(6).SubItems(2), ",", ".") & "'" & _
               " WHERE MUESTRA_ID = " & gmuestra & _
               "   AND ORDEN = 1 AND CAMPO_ID = 2812"
    If chkDuplicada.Value = Checked Then
        execute_bd "UPDATE SELLANTES_DETERMINACIONES " & _
                   "   SET VALOR_2 = '" & Replace(auxdatos.ListItems(12).SubItems(2), ",", ".") & "'" & _
                   " WHERE MUESTRA_ID = " & gmuestra & _
                   "   AND ORDEN = 1 AND CAMPO_ID = 2812"
    End If
End Sub

Private Sub Form_Load()
    log (Me.Name)
    cargar_botones Me
    cabecera
    cargar_combos
    permisos
    If gmuestra > 0 Then
        cargar_muestra
    End If
End Sub
Public Sub cabecera()
    With lista.ColumnHeaders
        .Add , , "", 300, lvwColumnLeft
        .Add , , "Ensayo", 2400, lvwColumnLeft
        .Add , , "R.Inferior", 1000, lvwColumnCenter
        .Add , , "R.Superior", 1000, lvwColumnCenter
        .Add , , "Resultado", 1200, lvwColumnRight
        .Add , , "Unidad", 1000, lvwColumnLeft
        .Add , , "Conforme", 1, lvwColumnCenter
        .Add , , "TIPO_DETERMINACION_ID", 1, lvwColumnCenter
        .Add , , "FORMULA_ID", 1, lvwColumnCenter
    End With
    With listaEquipos.ColumnHeaders
        .Add , , "NºEquipo", 800, lvwColumnLeft
        .Add , , "Nombre", 3200, lvwColumnLeft
        .Add , , "NºSerie", 1200, lvwColumnCenter
        .Add , , "Verificación", 1, lvwColumnCenter
    End With
    With listaReactivos.ColumnHeaders
        .Add , , "Número", 800, lvwColumnLeft
        .Add , , "Reactivo", 3200, lvwColumnLeft
        .Add , , "Caducidad", 1200, lvwColumnCenter
        .Add , , "TIPO", 0, lvwColumnCenter ' (I-E) Interno o externo
    End With
    ' Datos
    With datos.ColumnHeaders
        .Add , , "Campo", 3200, lvwColumnLeft
        .Add , , "Valor", 1500, lvwColumnRight
        .Add , , "Unidad", 1300, lvwColumnLeft
        .Add , , "ID", 0, lvwColumnCenter
        .Add , , "Enteros", 0, lvwColumnCenter
        .Add , , "Decimales", 0, lvwColumnCenter
    End With
    ' Aux Datos
    With auxdatos.ColumnHeaders
        .Add , , "MUESTRA_ID", 1, lvwColumnLeft
        .Add , , "ORDEN", 1, lvwColumnLeft
        .Add , , "Valor", 1000, lvwColumnLeft
        .Add , , "Linea", 1000, lvwColumnLeft
        .Add , , "Campo", 1000, lvwColumnLeft
        .Add , , "Media", 200, lvwColumnLeft
    End With
End Sub
Private Sub cargar_muestra()
    'Titulo
    Dim oMuestra As New clsMuestra
   On Error GoTo cargar_muestra_Error

    oMuestra.CargaMuestra (gmuestra)
    
    cargar_sellantes_cliente oMuestra.getCLIENTE_ID
    
    ' Duplicada
    If oMuestra.getANALISIS_DUPLICADO = 1 Then
        chkDuplicada.Value = Checked
    End If
    
    lbltitulo = "Registro resultados muestra : " & Trim(str(oMuestra.getID_GENERAL)) & " (" & oMuestra.CodigoParticular(gmuestra) & ")"
    Me.Caption = lbltitulo
    ' SE
    Dim oSe_Resultados As New clsSellantes_resultados
    oSe_Resultados.Carga (gmuestra)
    
    txtID_SELLANTE = oSe_Resultados.getSELLANTE_ID
    cmbProducto.BoundText = txtID_SELLANTE
    
    ' Cargar Ensayos
    llenar_combo cmbEnsayos, New clsSellantes_ensayos, oSe_Resultados.getSELLANTE_ID, Me, ""
    
    Dim oSellante As New clsSellantes
    With oSellante
        .Carga (oSe_Resultados.getSELLANTE_ID)
        txtDatos(0) = .getENSAYO
        txtDatos(1) = .getPROCESO
        txtDatos(2) = .getPREPARACION
        txtDatos(3) = .getPRODUCTO
    End With
    Dim oSe_Recepcion As New clsSellantes_recepcion
    With oSe_Recepcion
        .Carga (gmuestra)
        txtDatos(4) = .getN_MEZCLA
        txtDatos(12) = .getR_MEZCLA
        txtDatos(5) = .getLOTE
        txtDatos(6) = .getHIGROMETRIA
        txtDatos(7) = .getTEMPERATURA
'        If IsDate(.getFECHA) Then
'            fecha = Format(.getFECHA, "yyyy-mm-dd")
'        End If
        
        If IsDate(.getFECHA) Then
            If .getFECHA = "0000-00-00" Then
                chkfMezcla.Value = Checked
            Else
                chkfMezcla.Value = Unchecked
            End If
            fecha = Format(.getFECHA, "yyyy-mm-dd")
        Else
            fecha.Enabled = False
            chkfMezcla.Value = Checked
        End If
        
        If Trim(.getHORA) <> "" Then
            txtDatos(8) = Format(.getHORA, "hh:mm:ss")
        End If
        If IsDate(.getFECHA_LIMITE) Then
            If .getFECHA_LIMITE = "0000-00-00" Then
                chkFLimite.Value = Checked
            Else
                chkFLimite.Value = Unchecked
            End If
            fecha_limite = Format(.getFECHA_LIMITE, "yyyy-mm-dd")
        Else
            fecha_limite.Enabled = False
            chkFLimite.Value = Checked
'            chkFLimite.Visible = False
'            fecha_limite.Visible = False
'            lblCampos(0).Visible = False
        End If
        ' Recepcion
        txtDatos(11) = .getENSAYO_TEMPERATURA
        txtDatos(14) = .getENSAYO_HUMEDAD
        txtDatos(9) = .getENSAYO_ESPESOR
        ' Equipos
        If .getEQUIPOS <> "" Then
            cargar_equipos gmuestra
        End If
        ' Reactivos
        If .getREACTIVOS <> "" Or .getREACTIVOS_PROPIOS <> "" Then
            cargar_reactivos gmuestra
        End If
    End With
    ' Resultados
    cargar_resultados
    proteger_campos oMuestra.getCERRADA
    Set oSe_Resultados = Nothing
    lista_Click

   On Error GoTo 0
   Exit Sub

cargar_muestra_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure cargar_muestra of Formulario frmSE_Resultados"
End Sub

Private Sub lista_Click()
    If lista.ListItems.Count > 0 Then
        chkConforme(0).Value = False
        chkConforme(1).Value = False
        txtdato = lista.ListItems(lista.selectedItem.Index).SubItems(1)
        txtValor = lista.ListItems(lista.selectedItem.Index).SubItems(4)
        If lista.ListItems(lista.selectedItem.Index).SubItems(6) = "0" Then
            chkConforme(1).Value = True
        Else
            chkConforme(0).Value = True
        End If
'        If txtValor = "CONFORME" Then
'            chkConforme(0).value = True
'        End If
'        If txtValor = "NO CONFORME" Then
'            chkConforme(1).value = True
'        End If
        On Error Resume Next
        txtValor.SetFocus
        ' Criterio
        Dim oSe As New clsSellantes_ensayos
        oSe.Carga txtID_SELLANTE, lista.ListItems(lista.selectedItem.Index)
        txtDatos(10) = oSe.getCRITERIO
        ' Tipo determinacion
        datos.ListItems.Clear
        txtdato2 = ""
        txtvalor2 = ""
        If lista.ListItems(lista.selectedItem.Index).SubItems(7) <> 0 Then
            cargar_campos lista.ListItems(lista.selectedItem.Index).SubItems(7), lista.ListItems(lista.selectedItem.Index).SubItems(8)
        End If
    End If
End Sub

Private Sub listaEquipos_ItemCheck(ByVal Item As MSComctlLib.ListItem)
    grabar_equipos
End Sub

Private Sub txtdatos_GotFocus(Index As Integer)
    txtDatos(Index).BackColor = vbYellow
    txtDatos(Index).SelStart = 0
    txtDatos(Index).SelLength = Len(txtDatos(Index))
End Sub

Private Sub txtdatos_LostFocus(Index As Integer)
    txtDatos(Index).BackColor = vbWhite
    ' Unidades temperatura
    If Index = 7 Or Index = 11 Then
        If txtDatos(Index) <> "" Then
            txtDatos(Index) = Replace(txtDatos(Index), "ºC", "")
            txtDatos(Index) = txtDatos(Index) & "ºC"
        End If
    End If
    If Index = 6 Or Index = 14 Then
        If txtDatos(Index) <> "" Then
            txtDatos(Index) = Replace(txtDatos(Index), "%", "")
            txtDatos(Index) = txtDatos(Index) & "%"
        End If
    End If
End Sub

Private Sub txtvalor_GotFocus()
    txtValor.BackColor = vbYellow
    txtValor.SelStart = 0
    txtValor.SelLength = Len(txtValor)
End Sub

Private Sub txtvalor_KeyPress(KeyAscii As Integer)
    If txtdato = "" Then
        Exit Sub
    End If
    If KeyAscii = 46 Then
         KeyAscii = 44
    End If
    On Error GoTo fallo
    
    If KeyAscii = 13 Then
        KeyAscii = 0
        cmdAceptar_Click
    End If
    
    Exit Sub
fallo:
    error_grave "Error en frmListadoDeterminaciones(txtvalor_KeyPress) : " & Err.Description

End Sub

Private Sub txtvalor_LostFocus()
    txtValor.BackColor = vbWhite
End Sub
Public Sub colorear_linea(linea As Integer)
    If Trim(lista.ListItems(linea).SubItems(4)) <> "" And CInt(lista.ListItems(linea).SubItems(6)) = 0 Then
        lista.ListItems(linea).SmallIcon = 1
    Else
        lista.ListItems(linea).SmallIcon = 0
    End If
End Sub
Private Sub proteger_campos(CERRADA As Integer)
    If (CERRADA = 1 Or CERRADA = 3) And chkModificar.Value = Unchecked Then
        cmdAceptar.Enabled = False
        cmdModificar.Enabled = False
        cmdModificarSellante.Enabled = False
        
        cmdModificarEnsayo.Enabled = False
        cmdEliminarReactivo.Enabled = False
        cmdAnadirReactivo.Enabled = False
        cmdEliminarEquipo.Enabled = False
        cmdAnadirEquipo.Enabled = False
        cmbEquipos.desactivar
        cmbReactivos.desactivar
        Frame3.Enabled = False
        Frame5.Enabled = False
        cmbReactivosInternos.desactivar
        txtValor.Locked = True
        txtvalor2.Locked = True
        
        cmdEnsayoAdd.Enabled = False
        cmdEnsayoEliminar.Enabled = False
        cmbEnsayos.desactivar
    Else
        cmdAceptar.Enabled = True
        cmdModificar.Enabled = True
        cmdModificarSellante.Enabled = True
        
        cmdModificarEnsayo.Enabled = True
        cmdEliminarReactivo.Enabled = True
        cmdAnadirReactivo.Enabled = True
        cmdEliminarEquipo.Enabled = True
        cmdAnadirEquipo.Enabled = True
        cmbEquipos.activar
        cmbReactivos.activar
        cmbReactivosInternos.activar
        Frame3.Enabled = True
        Frame5.Enabled = True
        
        cmdEnsayoAdd.Enabled = True
        cmdEnsayoEliminar.Enabled = True
        cmbEnsayos.activar
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
End Sub

Private Sub cargar_equipos(muestra As Long)
'    Dim oSE As New clsSellantes_recepcion
'    With oSE
'    If .Carga(MUESTRA) Then
'            If .getEQUIPOS <> "" Then
'                Dim Equipos() As String
'
'                Dim oEquipo As New clsEquipos
'                Equipos = Split(.getEQUIPOS, ";")
'                For i = LBound(Equipos) To UBound(Equipos) - 1
'                    oEquipo.Carga_Datos_Basicos CLng(Equipos(i))
'                    With listaEquipos.ListItems.Add(, , oEquipo.getID_EQUIPO)
'                        .SubItems(1) = oEquipo.getNOMBRE
'                        .SubItems(2) = oEquipo.getSERIE
'                        .SubItems(3) = "0"
'                    End With
'                Next
'            End If
'    End If
'    End With
'    Set oSE = Nothing
    Dim oSe As New clsSellantes_equipos
    Dim rs As ADODB.Recordset
    Set rs = oSe.Listado(muestra)
    listaEquipos.ListItems.Clear
    If rs.RecordCount > 0 Then
        Do
            With listaEquipos.ListItems.Add(, , rs(0))
                .SubItems(1) = rs(1)
                .SubItems(2) = rs(2)
                .SubItems(3) = rs(5) ' VERIFICACION
            End With
            If rs("EN_INFORME") = 1 Then
                listaEquipos.ListItems(listaEquipos.ListItems.Count).Checked = True
            Else
                listaEquipos.ListItems(listaEquipos.ListItems.Count).Checked = False
            End If
            rs.MoveNext
        Loop Until rs.EOF
    End If
    Set oSe = Nothing
    
End Sub

Private Sub cargar_reactivos(muestra As Long)
    Dim oSe As New clsSellantes_recepcion
    With oSe
        If .Carga(muestra) Then
            ' REACTIVOS EXTERNOS
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
            If .getREACTIVOS_PROPIOS <> "" Then
                Dim REACTIVOS_PROPIOS() As String
                Dim oRPR As New clsRpr_botes
                Dim oTRPR As New clsRPR_Tipos
                REACTIVOS_PROPIOS = Split(.getREACTIVOS_PROPIOS, ";")
                For i = LBound(REACTIVOS_PROPIOS) To UBound(REACTIVOS_PROPIOS) - 1
                    oRPR.Carga CLng(REACTIVOS_PROPIOS(i))
                    oTRPR.CARGAR oRPR.getTIPO_REACTIVO_PR_ID
                    With listaReactivos.ListItems.Add(, , REACTIVOS_PROPIOS(i))
                        .SubItems(1) = oTRPR.getNOMBRE
                        .SubItems(2) = Format(oRPR.getFECHA_CADUCIDAD, "DD-MM-YYYY")
                        .SubItems(3) = "I"
                    End With
                Next
            End If
        End If
    End With
    Set oSe = Nothing
End Sub

Private Sub cargar_combos()
    llenar_combo cmbEquipos, New clsEquipos, 0, frmEquipoEdicion, ""
    llenar_combo cmbReactivos, New clsBotes_ex, 0, Me, " AND ABIERTO = 1 AND FINALIZADO = 0 "
    llenar_combo cmbReactivosInternos, New clsRpr_botes, 0, Me, " AND ISNULL(fecha_fin)"
End Sub
Private Sub cargar_campos(TIPO_DETERMINACION As Long, lFORMULA As Long)
    Dim ocampos As New clsFormulas_campos
    Dim rs As New ADODB.Recordset
    Dim consulta As String
    Dim duplicado As Integer
    Dim nombre As String
    Dim i As Integer
    Dim j As Integer
    cmdCalcular.Enabled = False
'    Dim oTD As New clsTipos_determinacion
'    oTD.CargarTipoDeterminacion TIPO_DETERMINACION
'    Set rs = ocampos.ListaFormulas(oTD.getFORMULA_ID)
    Set rs = ocampos.ListaFormulas(lFORMULA)
    lblestado.Caption = ""
    If chkDuplicada.Value = Checked Then
        duplicado = 2
        Label5(0).Width = 3900
        lblestado.Caption = "DUPLICADA"
        lblestado.visible = True
    Else
        duplicado = 1
        lblestado.visible = False
    End If
    Dim rs_campos As ADODB.Recordset
    Dim oSE_Deter As New clsSellantes_determinaciones
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
                    If oSE_Deter.Carga(gmuestra, lista.ListItems(lista.selectedItem.Index).Text, rs("id_campo")) = True Then
                      If j = 1 Then
                        .SubItems(1) = Replace(oSE_Deter.getVALOR_1, ".", ",")
                      Else
                        .SubItems(1) = Replace(oSE_Deter.getVALOR_2, ".", ",")
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
       With datos.ListItems.Add(, , "Resultado (Diferencia)")
          .SubItems(1) = " "
       End With
       With datos.ListItems.Add(, , "% Dif. entre duplicados")
          .SubItems(1) = " "
          .SubItems(2) = "%"
       End With
       With datos.ListItems.Add(, , "Desviación Estándar (1)")
           .SubItems(1) = " "
       End With
       With datos.ListItems.Add(, , "Desviación Estándar (2)")
           .SubItems(1) = " "
       End With
     Else
       With datos.ListItems.Add(, , "Desviación Estándar")
           .SubItems(1) = " "
       End With
     End If
     visualizar_duplicados
    End If
    ' Comprobar si ya tiene datos
    For i = 1 To auxdatos.ListItems.Count
        If lista.ListItems(lista.selectedItem.Index).Text = auxdatos.ListItems(i).SubItems(1) Then
            datos.ListItems(CInt(auxdatos.ListItems(i).SubItems(4))).SubItems(1) = auxdatos.ListItems(i).SubItems(2)
        End If
    Next
    Set rs = Nothing
    Set rs_campos = Nothing
    Set ocampos = Nothing
    datos_Click
End Sub
Private Sub visualizar_duplicados()
    ' Si la muestra es duplicada, visualizar resultados
    Dim numero_resultados As Integer
    Dim i As Integer
    Dim res1 As String
    Dim res2 As String
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
    Dim media As Single
    Dim dif As Single
    Dim dif_media As Single
    
    If numero_resultados = 2 And IsNumeric(res1) And IsNumeric(res2) Then ' Calcular media y diferencia
        ' JGM Datos de la diferencia
        dif = Abs((CSng(res1) - CSng(res2)))
        datos.ListItems(datos.ListItems.Count - 3).SubItems(1) = formatear(CStr(dif), 5, 2)
        ' Datos de la media
        media = (CSng(res1) + CSng(res2)) / 2
        datos.ListItems(datos.ListItems.Count - 4).SubItems(1) = formatear(CStr(media), 5, 2)
        ' Se modifica la diferencia para que siempre se muestre en %
        dif_media = (dif / media) * 100
        datos.ListItems(datos.ListItems.Count - 2).SubItems(1) = formatear(CStr(dif_media), 2, 2)
        If dif_media > 3.11 Then
            datos.ListItems(datos.ListItems.Count - 2).SubItems(2) = "NO CONFORME"
            datos.ListItems(datos.ListItems.Count - 2).ForeColor = vbRed
            datos.ListItems(datos.ListItems.Count - 2).bold = True
            datos.ListItems(datos.ListItems.Count - 2).ListSubItems(1).ForeColor = vbRed
            datos.ListItems(datos.ListItems.Count - 2).ListSubItems(2).ForeColor = vbRed
            
'            MsgBox "La diferencia de duplicados es mayor que la permitida (3,11 %)", vbExclamation, App.Title
        Else
            datos.ListItems(datos.ListItems.Count - 2).SubItems(2) = "CONFORME"
            datos.ListItems(datos.ListItems.Count - 2).ForeColor = vbBlack
            datos.ListItems(datos.ListItems.Count - 2).bold = False
            datos.ListItems(datos.ListItems.Count - 2).ListSubItems(1).ForeColor = vbBlack
            datos.ListItems(datos.ListItems.Count - 2).ListSubItems(2).ForeColor = vbBlack
        End If
        grabar_auxdatos
        
'        media = (CSng(res1) + CSng(res2)) / 2
'        datos.ListItems(datos.ListItems.Count - 3).SubItems(1) = Format(CStr(media), "##0.00")
'        grabar_auxdatos
'        dif = Abs((CSng(res1) - CSng(res2)))
'        datos.ListItems(datos.ListItems.Count - 2).SubItems(1) = Format(CStr(dif), "#,##0.00")
'        grabar_auxdatos
    Else
        If res1 = "--" Or res2 = "--" Then
            datos.ListItems(datos.ListItems.Count - 4).SubItems(1) = "--"
            datos.ListItems(datos.ListItems.Count - 3).SubItems(1) = "--"
            datos.ListItems(datos.ListItems.Count - 2).SubItems(1) = "--"
        Else
            If UCase(lblestado.Caption) = "DUPLICADA" Then
                datos.ListItems(datos.ListItems.Count - 4).SubItems(1) = lista.ListItems(lista.selectedItem.Index).SubItems(6)
            End If
        End If
    End If
    
    ' Desviación
    Dim sumatorio As Single
    Dim medida As Single
    Dim numero_medidas As Integer
    Dim resultado As Single
    media = 0
    sumatorio = 0
    numero_medidas = 0
    
    ' Primera medida
    If Trim(datos.ListItems(6).SubItems(1)) <> "" Then
        media = datos.ListItems(6).SubItems(1)
        For i = 1 To 5
            If IsNumeric(datos.ListItems(i).SubItems(1)) Then
                medida = datos.ListItems(i).SubItems(1)
                sumatorio = sumatorio + ((medida - media) * (medida - media))
                numero_medidas = numero_medidas + 1
            End If
        Next
        resultado = Sqr(sumatorio / (numero_medidas - 1))
        
        If chkDuplicada.Value = Unchecked Then
            datos.ListItems(7).SubItems(1) = formatear(CStr(resultado), 5, 3)
            If resultado > 0.677 Then
                datos.ListItems(7).SubItems(2) = "NO CONFORME"
                datos.ListItems(7).ForeColor = vbRed
                datos.ListItems(7).bold = True
                datos.ListItems(7).ListSubItems(1).ForeColor = vbRed
                datos.ListItems(7).ListSubItems(2).ForeColor = vbRed
            Else
                datos.ListItems(7).SubItems(2) = "CONFORME"
                datos.ListItems(7).ForeColor = vbBlack
                datos.ListItems(7).bold = False
                datos.ListItems(7).ListSubItems(1).ForeColor = vbBlack
                datos.ListItems(7).ListSubItems(2).ForeColor = vbBlack
            End If
        Else
            datos.ListItems(16).SubItems(1) = formatear(CStr(resultado), 5, 3)
            If resultado > 0.677 Then
                datos.ListItems(16).SubItems(2) = "NO CONFORME"
                datos.ListItems(16).ForeColor = vbRed
                datos.ListItems(16).bold = True
                datos.ListItems(16).ListSubItems(1).ForeColor = vbRed
                datos.ListItems(16).ListSubItems(2).ForeColor = vbRed
            Else
                datos.ListItems(16).SubItems(2) = "CONFORME"
                datos.ListItems(16).ForeColor = vbBlack
                datos.ListItems(16).bold = False
                datos.ListItems(16).ListSubItems(1).ForeColor = vbBlack
                datos.ListItems(16).ListSubItems(2).ForeColor = vbBlack
            End If
        End If
    End If
        ' Segunda medida para duplicados
    media = 0
    sumatorio = 0
    numero_medidas = 0
    If chkDuplicada.Value = Checked Then
        If Trim(datos.ListItems(12).SubItems(1)) <> "" Then
            media = datos.ListItems(12).SubItems(1)
            For i = 7 To 11
                If IsNumeric(datos.ListItems(i).SubItems(1)) Then
                    medida = datos.ListItems(i).SubItems(1)
                    sumatorio = sumatorio + ((medida - media) * (medida - media))
                    numero_medidas = numero_medidas + 1
                End If
            Next
            resultado = Sqr(sumatorio / (numero_medidas - 1))
            datos.ListItems(17).SubItems(1) = formatear(CStr(resultado), 5, 3)
            If resultado > 0.677 Then
                datos.ListItems(17).SubItems(2) = "NO CONFORME"
                datos.ListItems(17).ForeColor = vbRed
                datos.ListItems(17).bold = True
                datos.ListItems(17).ListSubItems(1).ForeColor = vbRed
                datos.ListItems(17).ListSubItems(2).ForeColor = vbRed
            Else
                datos.ListItems(17).SubItems(2) = "CONFORME"
                datos.ListItems(17).ForeColor = vbBlack
                datos.ListItems(17).bold = False
                datos.ListItems(17).ListSubItems(1).ForeColor = vbBlack
                datos.ListItems(17).ListSubItems(2).ForeColor = vbBlack
            End If
        End If
    End If
End Sub

Private Sub datos_Click()
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
        txtvalor2 = Trim(datos.ListItems(datos.selectedItem.Index).SubItems(1))
        txtvalor2.SetFocus
        txtvalor2.SelStart = 0
        txtvalor2.SelLength = Len(txtvalor2)
        txtdato2 = datos.ListItems(datos.selectedItem.Index)
    End If
End Sub

Private Sub grabar_auxdatos()
    Dim i As Integer
    For i = auxdatos.ListItems.Count To 1 Step -1
       If lista.ListItems(lista.selectedItem.Index).Text = auxdatos.ListItems(i).SubItems(1) Then
           auxdatos.ListItems.Remove (i)
       End If
    Next
    For i = 1 To datos.ListItems.Count
       With auxdatos.ListItems.Add(, , gmuestra) ' MUESTRA_ID
             .SubItems(1) = lista.ListItems(lista.selectedItem.Index).Text  ' ORDEN
             .SubItems(2) = datos.ListItems(i).SubItems(1) ' VALOR
             .SubItems(3) = i ' LINEA
             .SubItems(4) = datos.ListItems(i).SubItems(3) ' CAMPO
             If datos.ListItems(i).bold = True Then
                .bold = True
                ' Si es solucion, la subimoslas determinaciones
                If UCase(lblestado.Caption) <> "DUPLICADA" Then
                    If Trim(datos.ListItems(i).SubItems(1)) <> "" Then
'                        txtvalor = formatear(datos.ListItems(i).SubItems(1), 5, 1)
                        txtValor = formatear(datos.ListItems(i).SubItems(1), 5, 0)
                        insertar_resultado True
                    End If
                End If
             Else
                If UCase(lblestado.Caption) = "DUPLICADA" Then
                    If datos.ListItems(i).Text = "Resultado (MEDIA)" Then
                        .SubItems(5) = "M"
                    End If
                    If Trim(datos.ListItems(datos.ListItems.Count - 4).SubItems(1)) <> "" Then
'                        txtvalor = formatear(datos.ListItems(datos.ListItems.Count - 4).SubItems(1), 5, 1)
                        txtValor = formatear(datos.ListItems(datos.ListItems.Count - 4).SubItems(1), 5, 0)
                        insertar_resultado True
                    End If
                End If
             End If
       End With
    Next
    ' Grabar el resultado en la BD
    ' Ordenamos auxdatos por el CAMPO_ID
    If UCase(lblCerrada) <> "CERRADA" Then
        If chkDuplicada.Value = Checked Then
            auxdatos.Sorted = True
            auxdatos.SortKey = 4
        End If
    
        Dim oSE_Deter As New clsSellantes_determinaciones
        With oSE_Deter
            .Eliminar gmuestra, lista.ListItems(lista.selectedItem.Index).Text
    
            For i = 1 To auxdatos.ListItems.Count
                If auxdatos.ListItems(i).SubItems(4) <> "" Then ' Para la media y diferencia de duplicados
                    .setMUESTRA_ID = gmuestra
                    .setORDEN = auxdatos.ListItems(i).SubItems(1)
                    .setCAMPO_ID = auxdatos.ListItems(i).SubItems(4)
                    .setVALOR_1 = " "
                    If Trim(auxdatos.ListItems(i).SubItems(2)) <> "" Then
                        .setVALOR_1 = Replace(auxdatos.ListItems(i).SubItems(2), ",", ".")
                    End If
                    ' Valor duplicado
                    .setVALOR_2 = " "
                    If chkDuplicada.Value = Checked Then
                        i = i + 1
                        If Trim(auxdatos.ListItems(i).SubItems(2)) <> "" Then
                           .setVALOR_2 = Replace(auxdatos.ListItems(i).SubItems(2), ",", ".")
                        End If
                    End If
                    .Insertar
                End If
            Next
        End With
        Set oSE_Deter = Nothing
    End If
End Sub
Private Sub txtvalor2_GotFocus()
    txtvalor2.BackColor = vbYellow
    txtvalor2.SelStart = 0
    txtvalor2.SelLength = Len(txtvalor2)
End Sub

Private Sub txtvalor2_KeyPress(KeyAscii As Integer)
    If txtdato2 = "" Then
        Exit Sub
    End If
    If KeyAscii = 46 Then
         KeyAscii = 44
    End If
    On Error GoTo fallo
    
    If KeyAscii = 13 Then
        KeyAscii = 0
        If Trim(txtvalor2) = "" Or Trim(datos.ListItems(datos.selectedItem.Index).SubItems(3)) = "" Then
            datos.ListItems(datos.selectedItem.Index).SubItems(1) = " "
        Else
            datos.ListItems(datos.selectedItem.Index).SubItems(1) = formatear(txtvalor2, datos.ListItems(datos.selectedItem.Index).SubItems(5), datos.ListItems(datos.selectedItem.Index).SubItems(5))
        End If
        grabar_auxdatos
        visualizar_duplicados
        pasar_siguiente_campo
    End If
    
    Exit Sub
fallo:
    error_grave "Error en frmListadoDeterminaciones(txtvalor2_KeyPress) : " & Err.Description
End Sub
Private Sub pasar_siguiente_campo()
    If datos.ListItems.Count > datos.selectedItem.Index Then
        Set datos.selectedItem = datos.ListItems(datos.selectedItem.Index + 1)
        datos_Click
    Else
        If lista.ListItems.Count > lista.selectedItem.Index Then
            Set lista.selectedItem = lista.ListItems(lista.selectedItem.Index + 1)
            lista_Click
            datos_Click
        Else
            txtdato2 = ""
            txtvalor2 = ""
            datos2.SetFocus
        End If
    End If
End Sub

Private Sub txtvalor2_LostFocus()
    txtvalor2.BackColor = vbWhite
End Sub

Private Sub grabar_equipos()
    Dim Equipos As String
    Dim oSE_Equipos As New clsSellantes_equipos
    oSE_Equipos.Eliminar gmuestra
    Dim i As Integer
    For i = 1 To listaEquipos.ListItems.Count
        Equipos = Equipos & listaEquipos.ListItems(i).Text & ";"
        With oSE_Equipos
            .setMUESTRA_ID = gmuestra
            .setORDEN = i
            .setEQUIPO_ID = listaEquipos.ListItems(i).Text
            .setVERIFICACION_ID = listaEquipos.ListItems(i).SubItems(3)
            .setEN_INFORME = Abs(listaEquipos.ListItems(i).Checked)
            .Insertar
        End With
    Next
    ' Usos de los equipos
    Dim oEU As New clsEq_usos
    oEU.Eliminar gmuestra, 0
    For i = 1 To listaEquipos.ListItems.Count
      With oEU
          .setMUESTRA_ID = gmuestra
          .setEQUIPO_ID = listaEquipos.ListItems(i).Text
          .setDETERMINACION_ID = 0
          .setUSOS = 1
          .Insertar
      End With
    Next
    Set oEU = Nothing
    ' Recepcion
    Dim oSe As New clsSellantes_recepcion
    oSe.ModificarEquipos gmuestra, Equipos
    Set oSe = Nothing
End Sub
Private Sub grabar_reactivos()
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
    Dim oSe As New clsSellantes_recepcion
    oSe.ModificarReactivos gmuestra, Reactivo, REACTIVOS_PROPIOS
    Set oSe = Nothing
End Sub

Private Sub cargar_resultados()
    Dim rs As ADODB.Recordset
    Dim oSe_Resultados As New clsSellantes_resultados
    Set rs = oSe_Resultados.Listado_Resultados(gmuestra)
    lista.ListItems.Clear
    If rs.RecordCount > 0 Then
        Do
            With lista.ListItems.Add(, , rs(0))
              .SubItems(1) = rs(1)
              .SubItems(2) = rs(2)
              .SubItems(3) = rs(3)
              If Trim(rs(4)) = "" Then
                  .SubItems(4) = " "
              Else
                  .SubItems(4) = Trim(rs(4))
              End If
              .SubItems(5) = rs(5)
              .SubItems(6) = rs(6)
              .SubItems(7) = rs(8) ' TIPO_DETERMINACION_ID
              .SubItems(8) = rs(12) ' FORMULA_ID
            End With
            colorear_linea (lista.ListItems.Count)
            rs.MoveNext
        Loop Until rs.EOF
    End If

End Sub

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
Private Sub cargar_sellantes_cliente(CLIENTE_ID As Long)
    Dim oSellante As New clsSellantes
    Set cmbProducto.RowSource = oSellante.Listado_Combo_Sellantes_de_Clientes(CLIENTE_ID)
    cmbProducto.ListField = "C2"
    cmbProducto.DataField = "C1" 'campo asociado
    cmbProducto.BoundColumn = "C1" 'lo que realmente
    Set oSellante = Nothing
End Sub

