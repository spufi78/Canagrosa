VERSION 5.00
Object = "{0E59F1D2-1FBE-11D0-8FF2-00A0D10038BC}#1.0#0"; "msscript.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{CEA3B7F6-2847-4E5E-A551-DB7A62489D44}#46.0#0"; "miCombo.ocx"
Begin VB.Form frmCE_Listado_Probetas 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Listado de Probetas pendientes de Analizar"
   ClientHeight    =   9000
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   13620
   Icon            =   "frmCE_Listado_Probetas.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   9000
   ScaleWidth      =   13620
   Begin Geslab.ControlPanelXP cpOpciones 
      Height          =   2130
      Left            =   7380
      TabIndex        =   34
      Top             =   1620
      Width           =   6180
      _ExtentX        =   10901
      _ExtentY        =   3757
      Caption         =   "Ensayos de Tiempo"
      BackColor       =   12632256
      TextColor       =   0
      HeaderColor     =   8421504
      PanelOpen       =   0   'False
      Object.Height          =   2130
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
         Height          =   1410
         Left            =   135
         TabIndex        =   35
         Top             =   495
         Width           =   5910
         Begin VB.CommandButton cmdComienzo 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Comenzar Marcados"
            Height          =   915
            Left            =   3915
            Picture         =   "frmCE_Listado_Probetas.frx":1272
            Style           =   1  'Graphical
            TabIndex        =   36
            Top             =   270
            Width           =   1725
         End
         Begin MSComCtl2.DTPicker ddesde 
            Height          =   330
            Left            =   1305
            TabIndex        =   37
            Top             =   360
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
            Format          =   51314689
            CurrentDate     =   38000
            MinDate         =   2
         End
         Begin MSComCtl2.DTPicker dhdesde 
            Height          =   330
            Left            =   2655
            TabIndex        =   38
            Top             =   360
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
            Format          =   51314690
            CurrentDate     =   38000
            MinDate         =   2
         End
         Begin MSComCtl2.DTPicker dhasta 
            Height          =   330
            Left            =   1305
            TabIndex        =   39
            Top             =   765
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
            Format          =   51314689
            CurrentDate     =   38000
            MinDate         =   2
         End
         Begin MSComCtl2.DTPicker dhhasta 
            Height          =   330
            Left            =   2655
            TabIndex        =   40
            Top             =   765
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
            Format          =   51314690
            CurrentDate     =   38000
            MinDate         =   2
         End
         Begin VB.Label lblCampos 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C0C0&
            Caption         =   "Fecha de fin"
            Height          =   195
            Index           =   2
            Left            =   90
            TabIndex        =   42
            Top             =   810
            Width           =   885
         End
         Begin VB.Label lblCampos 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C0C0&
            Caption         =   "Fecha de inicio"
            Height          =   195
            Index           =   4
            Left            =   90
            TabIndex        =   41
            Top             =   450
            Width           =   1080
         End
      End
   End
   Begin Geslab.ControlPanelXP cpFormula 
      Height          =   5865
      Left            =   7380
      TabIndex        =   13
      Top             =   2115
      Width           =   6180
      _ExtentX        =   10901
      _ExtentY        =   10345
      Caption         =   "Formula"
      BackColor       =   12632256
      TextColor       =   0
      HeaderColor     =   8421504
      PanelOpen       =   0   'False
      Object.Height          =   5865
      Begin VB.Frame Frame3 
         BackColor       =   &H00C0C0C0&
         Height          =   720
         Left            =   90
         TabIndex        =   15
         Top             =   5040
         Width           =   5955
         Begin VB.CheckBox chkDuplicada 
            Caption         =   "Duplicada"
            Height          =   195
            Left            =   0
            TabIndex        =   53
            Top             =   0
            Visible         =   0   'False
            Width           =   1050
         End
         Begin VB.TextBox txtdato 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            Height          =   330
            Left            =   765
            Locked          =   -1  'True
            TabIndex        =   18
            Top             =   225
            Width           =   2310
         End
         Begin VB.TextBox txtvalor 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            Height          =   330
            Left            =   3645
            TabIndex        =   17
            Top             =   225
            Width           =   1635
         End
         Begin VB.CommandButton cmdcalcular 
            BackColor       =   &H00E0E0E0&
            Enabled         =   0   'False
            Height          =   555
            Left            =   5355
            Picture         =   "frmCE_Listado_Probetas.frx":1B3C
            Style           =   1  'Graphical
            TabIndex        =   16
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
            Index           =   0
            Left            =   90
            TabIndex        =   20
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
            Index           =   3
            Left            =   3150
            TabIndex        =   19
            Top             =   270
            Width           =   555
         End
      End
      Begin MSComctlLib.ListView datos 
         Height          =   4545
         Left            =   45
         TabIndex        =   14
         Top             =   450
         Width           =   5985
         _ExtentX        =   10557
         _ExtentY        =   8017
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
      Begin VB.Label lblestado 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
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
         Left            =   3060
         TabIndex        =   52
         Top             =   45
         Visible         =   0   'False
         Width           =   2715
      End
   End
   Begin VB.CommandButton cmdmodificarprobetas 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Datos Probetas"
      Height          =   825
      Left            =   5580
      Picture         =   "frmCE_Listado_Probetas.frx":1E46
      Style           =   1  'Graphical
      TabIndex        =   33
      Top             =   8100
      Width           =   1095
   End
   Begin VB.CommandButton cmdImagen 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Imagenes"
      Height          =   825
      Left            =   2205
      Style           =   1  'Graphical
      TabIndex        =   32
      Top             =   8100
      Width           =   1095
   End
   Begin VB.CommandButton cmdPNT 
      BackColor       =   &H00E0E0E0&
      Caption         =   "P.N.T."
      Height          =   825
      Left            =   3330
      Picture         =   "frmCE_Listado_Probetas.frx":2694
      Style           =   1  'Graphical
      TabIndex        =   31
      Top             =   8100
      Width           =   1095
   End
   Begin VB.CommandButton cmdtipoensayo 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Tipo Ensayo"
      Height          =   825
      Left            =   4455
      Picture         =   "frmCE_Listado_Probetas.frx":2F5E
      Style           =   1  'Graphical
      TabIndex        =   30
      Top             =   8100
      Width           =   1095
   End
   Begin Geslab.ControlPanelXP cpConforme 
      Height          =   3660
      Left            =   7380
      TabIndex        =   21
      Top             =   3780
      Width           =   6180
      _ExtentX        =   10901
      _ExtentY        =   6456
      Caption         =   "Resultado Conforme/No Conforme"
      BackColor       =   12632256
      TextColor       =   0
      HeaderColor     =   8421504
      PanelOpen       =   0   'False
      Object.Height          =   3660
      Begin VB.TextBox txtRangoMin 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         Height          =   285
         Left            =   5040
         Locked          =   -1  'True
         TabIndex        =   46
         Top             =   765
         Width           =   1020
      End
      Begin VB.TextBox txtRangoMax 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         Height          =   285
         Left            =   5040
         Locked          =   -1  'True
         TabIndex        =   45
         Top             =   1125
         Width           =   1020
      End
      Begin VB.TextBox txtcriterio 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         Height          =   1455
         Left            =   135
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   28
         Top             =   810
         Width           =   3870
      End
      Begin VB.Frame Frame2 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         ForeColor       =   &H80000008&
         Height          =   1110
         Left            =   135
         TabIndex        =   22
         Top             =   2340
         Width           =   5940
         Begin VB.TextBox txtconformeresultado 
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
            Left            =   990
            TabIndex        =   26
            Top             =   630
            Width           =   2025
         End
         Begin VB.CommandButton cmdAceptarConforme 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Aceptar"
            Height          =   840
            Left            =   4950
            Picture         =   "frmCE_Listado_Probetas.frx":3828
            Style           =   1  'Graphical
            TabIndex        =   25
            Top             =   180
            Width           =   915
         End
         Begin VB.OptionButton chkConforme 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            Caption         =   "Conforme"
            ForeColor       =   &H80000008&
            Height          =   195
            Index           =   0
            Left            =   3375
            TabIndex        =   24
            Top             =   315
            Width           =   1050
         End
         Begin VB.OptionButton chkConforme 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            Caption         =   "NO Conforme"
            ForeColor       =   &H80000008&
            Height          =   195
            Index           =   1
            Left            =   3375
            TabIndex        =   23
            Top             =   720
            Width           =   1275
         End
         Begin MSComCtl2.DTPicker fechaconforme 
            Height          =   330
            Left            =   990
            TabIndex        =   44
            Top             =   225
            Width           =   1245
            _ExtentX        =   2196
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
         Begin VB.Label Label1 
            BackColor       =   &H00C0C0C0&
            Caption         =   "Fecha"
            ForeColor       =   &H00000000&
            Height          =   255
            Index           =   4
            Left            =   135
            TabIndex        =   43
            Top             =   315
            Width           =   795
         End
         Begin VB.Label Label1 
            BackColor       =   &H00C0C0C0&
            Caption         =   "Resultado"
            ForeColor       =   &H00000000&
            Height          =   255
            Index           =   1
            Left            =   135
            TabIndex        =   27
            Top             =   720
            Width           =   795
         End
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Rango Min."
         Height          =   195
         Index           =   14
         Left            =   4140
         TabIndex        =   48
         Top             =   810
         Width           =   825
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Rango Max."
         Height          =   195
         Index           =   15
         Left            =   4140
         TabIndex        =   47
         Top             =   1170
         Width           =   870
      End
      Begin VB.Label Label1 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Criterio de Aceptación"
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   2
         Left            =   135
         TabIndex        =   29
         Top             =   540
         Width           =   1740
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Datos de Búsqueda"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   45
      TabIndex        =   6
      Top             =   585
      Width           =   13545
      Begin VB.CheckBox chkNoIniciados 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Caption         =   "Mostrar sólo Iniciados"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   1
         Left            =   4590
         TabIndex        =   51
         Top             =   630
         Width           =   2220
      End
      Begin VB.CheckBox chkNoIniciados 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Caption         =   "Mostrar sólo No Iniciados"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   0
         Left            =   2385
         TabIndex        =   50
         Top             =   630
         Width           =   2220
      End
      Begin VB.CheckBox chkEnsayosTiempo 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Caption         =   "Sólo ensayos de Tiempo"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   135
         TabIndex        =   49
         Top             =   630
         Width           =   2220
      End
      Begin VB.CommandButton cmdBuscar 
         BackColor       =   &H00E0E0E0&
         Caption         =   "&Buscar"
         Height          =   780
         Left            =   12600
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   135
         Width           =   885
      End
      Begin VB.CheckBox chkTodas 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Caption         =   "Todas"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   11025
         TabIndex        =   7
         Top             =   315
         Width           =   960
      End
      Begin pryCombo.miCombo cmbCE 
         Height          =   330
         Left            =   1440
         TabIndex        =   8
         Top             =   270
         Width           =   9465
         _ExtentX        =   16695
         _ExtentY        =   582
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Ensayo Eficacia"
         Height          =   195
         Index           =   3
         Left            =   135
         TabIndex        =   10
         Top             =   315
         Width           =   1155
      End
   End
   Begin VB.CommandButton cmdVerMuestra 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Muestra"
      Height          =   825
      Left            =   45
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   8100
      Width           =   1050
   End
   Begin VB.CommandButton cmdcancel 
      BackColor       =   &H00E0E0E0&
      Caption         =   "ESC-Salir"
      Height          =   870
      Left            =   12540
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   8070
      Width           =   1050
   End
   Begin MSComctlLib.ListView auxdatos 
      Height          =   3015
      Left            =   315
      TabIndex        =   1
      Top             =   4770
      Visible         =   0   'False
      Width           =   6675
      _ExtentX        =   11774
      _ExtentY        =   5318
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
   Begin VB.CommandButton cmdDeter 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Registro"
      Height          =   825
      Left            =   1125
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   8100
      Width           =   1050
   End
   Begin MSComctlLib.ListView lista 
      Height          =   6435
      Left            =   45
      TabIndex        =   0
      Top             =   1575
      Width           =   1500
      _ExtentX        =   2646
      _ExtentY        =   11351
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      AllowReorder    =   -1  'True
      Checkboxes      =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   8388608
      BackColor       =   13621491
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
   Begin MSScriptControlCtl.ScriptControl sc 
      Left            =   9450
      Top             =   4395
      _ExtentX        =   1005
      _ExtentY        =   1005
   End
   Begin MSComctlLib.ListView probetas 
      Height          =   6435
      Left            =   1530
      TabIndex        =   11
      Top             =   1575
      Width           =   5820
      _ExtentX        =   10266
      _ExtentY        =   11351
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
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   10530
      Top             =   8055
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   3
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCE_Listado_Probetas.frx":40F2
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCE_Listado_Probetas.frx":49CC
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCE_Listado_Probetas.frx":52A6
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Label lbltitulo 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Caption         =   "Control de Probetas Pendientes de Analizar"
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
      Left            =   180
      TabIndex        =   12
      Top             =   135
      Width           =   11505
   End
   Begin VB.Label lbltotal 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
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
      ForeColor       =   &H000000C0&
      Height          =   195
      Left            =   2295
      TabIndex        =   5
      Top             =   8100
      Width           =   3045
   End
   Begin VB.Shape fondo 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   555
      Left            =   0
      Top             =   0
      Width           =   13590
   End
End
Attribute VB_Name = "frmCE_Listado_Probetas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub chkEnsayosTiempo_Click()
    cmdBuscar_Click
End Sub
Private Sub chkNoIniciados_Click(Index As Integer)
    If Index = 0 Then
        If chkNoIniciados(0).Value = Checked Then
            chkNoIniciados(1).Value = Unchecked
        End If
    Else
        If chkNoIniciados(1).Value = Checked Then
            chkNoIniciados(0).Value = Unchecked
        End If
    End If
    cmdBuscar_Click
End Sub

Private Sub chkTodas_Click()
    txtp1 = ""
    txtp2 = ""
    If chkTodas.Value = Checked Then
        cmbCE.limpiar
        cmbCE.desactivar
    Else
        cmbCE.activar
    End If
End Sub
Private Sub cmbTiposMuestra_change()
    If cmbTiposMuestra.getTEXTO <> "" Then
        Me.MousePointer = 11
        cmbDeter.Text = ""
        cmbDeter.Enabled = False
        Dim consulta As String
        Dim rs As New ADODB.Recordset
        consulta = "SELECT distinct id_tipo_determinacion, CONCAT(td.nombre,' ',td.descripcion) as a" & _
            " FROM muestras mu, determinaciones de, tipos_determinacion td" & _
            " WHERE mu.tipo_muestra_id=" & cmbTiposMuestra.getPK_SALIDA & _
            "  AND mu.anno = " & txtanno & _
            "  AND mu.ANULADA = 0 " & _
            "  AND de.tipo_determinacion_id=id_tipo_determinacion" & _
            "  AND de.muestra_id=id_muestra" & _
            "  AND (de.resultado IS NULL OR de.resultado = '') " & _
            " order by td.nombre"
        Set rs = datos_bd(consulta)
        Set cmbDeter.RowSource = rs
        cmbDeter.ListField = "a"   'lo que enseña
        cmbDeter.DataField = "id_tipo_determinacion" 'campo asociado
        cmbDeter.BoundColumn = "id_tipo_determinacion" 'lo que realmente envia
        Set rs = Nothing
        Me.MousePointer = 0
        cmbDeter.Enabled = True
    End If
End Sub

Private Sub cmbCE_change()
    cmdBuscar_Click
End Sub

Private Sub cmdAceptarConforme_Click()
    Dim oCe_resultados As New clsCe_resultados
    With oCe_resultados
       .setFECHA = Format(fechaconforme, "dd/mm/yyyy")
       .setRESULTADO = txtconformeresultado
       ' Conforme/No conforme
       If chkConforme(0).Value = True And txtconformeresultado = "" Then
           .setCONFORME = 1
       Else
         If Trim(txtconformeresultado) <> "" Then
           If IsNumeric(txtconformeresultado) Then
             .setCONFORME = 1
             If Trim(txtRangoMin) <> "" And IsNumeric(txtRangoMin) Then
               If CSng(Replace(txtRangoMin, ".", ",")) > CSng(Replace(txtconformeresultado, ".", ",")) Then
                 .setCONFORME = 0
               End If
             End If
             If Trim(txtRangoMax) <> "" And IsNumeric(txtRangoMax) Then
               If CSng(Replace(txtRangoMax, ".", ",")) < CSng(Replace(txtconformeresultado, ".", ",")) Then
                 .setCONFORME = 0
               End If
             End If
           Else
             If chkConforme(0).Value = True Then
                .setCONFORME = 1
             Else
                .setCONFORME = 0
             End If
           End If
         Else
           .setCONFORME = 0
         End If
       End If
       Dim lMUESTRA As Long
       Dim lDESIGNACION As String
       Dim lPROBETA As Integer
       Dim lAREA As Integer
       lMUESTRA = CLng(probetas.ListItems(probetas.selectedItem.Index).Text)
       lDESIGNACION = probetas.ListItems(probetas.selectedItem.Index).SubItems(1)
       lPROBETA = CInt(probetas.ListItems(probetas.selectedItem.Index).SubItems(2))
       lAREA = CInt(probetas.ListItems(probetas.selectedItem.Index).SubItems(3))
       .Modificar_Resultado lMUESTRA, lDESIGNACION, lPROBETA, lAREA, True
    End With
    Set oCe_resultados = Nothing
End Sub

Private Sub cmdcancel_Click()
'    If MsgBox("Los datos no guardados se perderan. ¿Desea salir?", vbQuestion + vbYesNo, App.Title) = vbYes Then
      Unload Me
'    End If
End Sub

Private Sub cmdComienzo_Click()
    Dim s As String
   On Error GoTo cmdComienzo_Click_Error

    s = "¿Establecer fechas para los ensayos marcados?  Se generará un aviso de inicio y fin."
    If MsgBox(s, vbQuestion + vbYesNo, App.Title) = vbYes Then
        Dim oMensaje As New clsMensajes
        Dim oMuestra As New clsMuestra
        Dim oce_recepcion As New clsCe_recepcion
        Dim mens As Long
        Dim i As Integer
        Dim j As Integer
        ddesde = Date
        dhdesde = Date & " " & Time
        ' Carga de usuarios para envio de mensaje
        Dim omu As New clsMensajes_usuarios
        Dim usuarios() As String
        Dim opar As New clsParametros
        If (opar.Carga(11, "")) Then
            usuarios = Split(opar.getVALOR, ",")
        End If
        ' Enviar aviso
        For i = 1 To lista.ListItems.Count
            If lista.ListItems(i).Checked = True And lista.ListItems(i).SubItems(6) <> "" Then
                ' Calcular fecha de terminacion
'                If lista.ListItems(i).SubItems(6) <> "" Then
                    dhhasta = DateAdd("h", lista.ListItems(i).SubItems(6), dhdesde)
                    dhasta = dhhasta
'                End If
                oMuestra.CargaMuestra (lista.ListItems(i).SubItems(1))
                With oMensaje
                    .setASUNTO = Trim(str(oMuestra.getID_GENERAL)) & " (" & oMuestra.CodigoParticular(lista.ListItems(i).SubItems(1)) & ")" & " Finalización de Control de eficacia"
                    .setTEXTO = ""
                    .setTEXTO = .getTEXTO & "El usuario " & USUARIO.getUSUARIO & " ha iniciado un control de eficacia. " & vbNewLine & vbNewLine
                    .setTEXTO = .getTEXTO & "Fecha de comienzo : " & dhdesde & vbNewLine & vbNewLine
                    .setTEXTO = .getTEXTO & "Fecha de finalización : " & dhhasta & vbNewLine
                    .setEMPLEADO_ID = USUARIO.getID_EMPLEADO
                    
                    .setFECHA_INICIO = Format(dhhasta.Value, "yyyy-mm-dd")
                    .setFECHA_FIN = Format(dhhasta.Value, "yyyy-mm-dd")
                    
                    .setACCION = "frmVerMuestra;" & PK_ID_MUESTRA
                    .setHORA_INICIO = Format(dhhasta.Value, "hh:mm:ss")
                    .setHORA_FIN = Format(dhhasta.Value, "hh:mm:ss")
                    .setCATEGORIA = MENSAJES_CATEGORIAS.MENSAJES_CATEGORIAS_CE
                    .setDURACION = 0
                    
                    mens = .Insertar
                    If mens > 0 Then
                        For j = LBound(usuarios) To UBound(usuarios)
                            If usuarios(j) <> "" Then
                                omu.setEMPLEADO_ID = usuarios(j)
                                omu.setMENSAJE_ID = mens
                                omu.Insertar
                            End If
                        Next
                    End If
                End With
                With oce_recepcion
                    .setDURACION_FECHA_DESDE = Format(ddesde.Value, "dd-mm-yyyy")
                    .setDURACION_HORA_DESDE = Format(dhdesde.Value, "hh:mm:ss")
                    .setDURACION_FECHA_HASTA = Format(dhhasta.Value, "dd-mm-yyyy")
                    .setDURACION_HORA_HASTA = Format(dhhasta.Value, "hh:mm:ss")
                    .Informar_Duracion_Ensayo lista.ListItems(i).SubItems(1)
                End With
            End If
        Next
        frmCalendario.cargar_eventos
        MsgBox "Fechas establecidas correctamente.", vbInformation, App.Title
        cmdBuscar_Click
    End If

   On Error GoTo 0
   Exit Sub

cmdComienzo_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure cmdComienzo_Click of Formulario frmCE_Resultados2"

End Sub

Private Sub cmdDeter_Click()
    If lista.ListItems.Count > 0 Then
        frmCE_Resultados.PK_ID_MUESTRA = lista.ListItems(lista.selectedItem.Index).SubItems(1)
        frmCE_Resultados.Show 1
    End If
End Sub

Private Sub cmdImagen_Click()
    If lista.ListItems.Count > 0 Then
        With frmCE_Imagenes
            .PK = lista.ListItems(lista.selectedItem.Index).SubItems(1)
            .Show 1
        End With
    End If
End Sub

Private Sub cmdmodificarprobetas_Click()
    If lista.ListItems.Count > 0 Then
        With frmCE_Recepcion_Probetas
            .PK_MUESTRA = lista.ListItems(lista.selectedItem.Index).SubItems(1)
            .Show 1
        End With
    End If
End Sub

Private Sub cmdPNT_Click()
    If lista.ListItems.Count > 0 Then
        Dim oce_recepcion As New clsCe_recepcion
        oce_recepcion.Carga lista.ListItems(lista.selectedItem.Index).SubItems(1)
        Dim oCE As New clsCe_tipos_ensayos
        oCE.Carga CLng(oce_recepcion.getTIPO_ENSAYO_ID)
        If oCE.getPNT_VINCULADO <> 0 Then
            Dim oPNT As New clsCa_documentos
            oPNT.mostrar oCE.getPNT_VINCULADO, False
            Set oPNT = Nothing
        Else
            MsgBox "El Tipo de Ensayo no tiene PNT Vínculado.", vbExclamation, App.Title
        End If
        Set oce_recepcion = Nothing
        Set oCE = Nothing
    End If

End Sub

Private Sub cmdtipoensayo_Click()
    If lista.ListItems.Count > 0 Then
        Dim oce_recepcion As New clsCe_recepcion
        oce_recepcion.Carga lista.ListItems(lista.selectedItem.Index).SubItems(1)
        frmCE_Tipo_Ensayo.PK = CLng(oce_recepcion.getTIPO_ENSAYO_ID)
        frmCE_Tipo_Ensayo.Show 1
        Set oce_recepcion = Nothing
    End If
End Sub

Private Sub cmdVerMuestra_Click()
    If lista.ListItems.Count > 0 Then
        gmuestra = lista.ListItems(lista.selectedItem.Index).SubItems(1)
        frmVerMuestra.Show 1
    End If
End Sub
Private Sub Form_Load()
    log (Me.Name)
    cargar_botones Me
    Me.Left = 50
    Me.top = 50
    cabecera
    cargar_ensayos
End Sub
Private Sub cabecera()
    ' Cabecera de las muestras
    With lista.ColumnHeaders
        .Add , , "Código", 1250, lvwColumnLeft
        .Add , , "ID", 1, lvwColumnLeft
        .Add , , "FORMULA", 1, lvwColumnLeft
        .Add , , "CRITERIO", 1, lvwColumnLeft
        .Add , , "RANGOMIN", 1, lvwColumnLeft
        .Add , , "RANGOMAX", 1, lvwColumnLeft
        .Add , , "HORAS", 1, lvwColumnLeft
        .Add , , "FINICIO", 1, lvwColumnLeft
        .Add , , "HINICIO", 1, lvwColumnLeft
        .Add , , "FFIN", 1, lvwColumnLeft
        .Add , , "HFIN", 1, lvwColumnLeft
        .Add , , "DUPLICADA", 1, lvwColumnLeft
    End With
    ' Probetas
    With probetas.ColumnHeaders
        .Add , , "", 0, lvwColumnLeft
        .Add , , "DESIGNACION", 1, lvwColumnLeft
        .Add , , "PROBETA", 1, lvwColumnLeft
        .Add , , "AREA", 1, lvwColumnLeft
        .Add , , "Identificación Canagrosa", 3000, lvwColumnLeft
        .Add , , "Fecha", 1100, lvwColumnCenter
        .Add , , "Resultado", 1000, lvwColumnRight
        .Add , , "Conforme", 1, lvwColumnRight
    End With
    ' Datos
    With datos.ColumnHeaders
        .Add , , "Campo", 3000, lvwColumnLeft
        .Add , , "Valor", 1500, lvwColumnRight
        .Add , , "Unidad", 1000, lvwColumnLeft
        .Add , , "ID", 0, lvwColumnCenter
        .Add , , "Enteros", 0, lvwColumnCenter
        .Add , , "Decimales", 0, lvwColumnCenter
    End With
    ' Aux Datos
    With auxdatos.ColumnHeaders
        .Add , , "DESIGNACION", 100, lvwColumnLeft
        .Add , , "PROBETA", 100, lvwColumnLeft
        .Add , , "AREA", 100, lvwColumnLeft
        .Add , , "Valor", 1000, lvwColumnLeft
        .Add , , "Linea", 1000, lvwColumnLeft
        .Add , , "Campo", 1000, lvwColumnLeft
        .Add , , "Media", 200, lvwColumnLeft
    End With
End Sub
Private Sub cargar_ensayos()
    Dim consulta As String
    consulta = "SELECT DISTINCT B.ID_TIPO_MUESTRA , B.NOMBRE " & _
               "  FROM MUESTRAS A, TIPOS_MUESTRA B " & _
               " WHERE A.ANALISIS_MODIFICADO = 2 " & _
               "   AND B.TIPO_ESPECIAL_ID IN (" & tipo_especial.control_eficacia & "," & tipo_especial.CONTROLES_PROCESOS & ")" & _
               "   AND A.TIPO_MUESTRA_ID = B.ID_TIPO_MUESTRA " & _
               "   AND A.CERRADA = 0 AND A.ANULADA = 0"
    Dim conn As ADODB.Connection
    If CrearConexionGlobal(conn, "", "") = True Then ' CONECTAR EL RS
        With cmbCE
            .setCONN = conn
            .setFK_CAMPO = ""
            .setFK_VALOR = 0
            .setTABLA = "TIPOS_MUESTRA"
            .setDESCRIPCION = "Tipos Ensayos de Eficacia"
            .setPK = "B.ID_TIPO_MUESTRA"
            .setCAMPO = "NOMBRE"
            .setMUESTRA_DETALLE = True
            .setQUERY = consulta
            .setFILTRO = ""
            Set .FORMULARIO = frmTM_Detalle
        End With
    End If
End Sub
Private Sub cmdBuscar_Click()
    buscar
End Sub
Private Sub buscar()
    Dim oCe_resultados As New clsCe_resultados
    Dim rs As ADODB.Recordset
    If cmbCE.getTEXTO = "" Then
        Set rs = oCe_resultados.Listado_Pendientes(0, chkEnsayosTiempo.Value, chkNoIniciados(0).Value, chkNoIniciados(1).Value)
    Else
        Set rs = oCe_resultados.Listado_Pendientes(cmbCE.getPK_SALIDA, chkEnsayosTiempo.Value, chkNoIniciados(0).Value, chkNoIniciados(1).Value)
    End If
    Dim i As Integer
    lista.ListItems.Clear
    probetas.ListItems.Clear
    If rs.RecordCount > 0 Then
       Do
           With lista.ListItems.Add(, , rs(1)) ' CODIGO
               .SubItems(1) = rs(0) ' ID_MUESTRA
               .SubItems(2) = rs(2) ' fORMULA
               .SubItems(3) = rs(3) ' Criterio
               .SubItems(4) = rs(4) ' RANGO MIN
               .SubItems(5) = rs(5) ' RANGO MAX
               .SubItems(6) = rs(6) ' HORAS
               .SubItems(7) = rs(7) ' F.Inicio
               .SubItems(8) = rs(8) ' H.Inicio
               .SubItems(9) = rs(9) ' F.Fin
               .SubItems(10) = rs(10) ' H.Fin
               .SubItems(11) = rs(11) ' Duplicada
           End With
           rs.MoveNext
       Loop Until rs.EOF
       lista_Click
    End If
    Set rs = Nothing
End Sub


Private Sub lista_Click()
    If lista.ListItems.Count = 0 Then
        Exit Sub
    End If
    auxdatos.ListItems.Clear
    datos.ListItems.Clear
    ' Por formula o resultado
    If CInt(lista.ListItems(lista.selectedItem.Index).SubItems(2)) = 0 Then
        cpFormula.PanelOpen = False
        cpFormula.CanExpand = False
        cpConforme.PanelOpen = True
        cpConforme.CanExpand = True
    Else
        ' Formula
        cpFormula.PanelOpen = True
        cpFormula.CanExpand = True
        cpConforme.PanelOpen = False
        cpConforme.CanExpand = False
        chkDuplicada.Value = lista.ListItems(lista.selectedItem.Index).SubItems(11)
    End If
    ' Ensayo de Tiempo
    If Trim(lista.ListItems(lista.selectedItem.Index).SubItems(6)) = "" Then
        cpOpciones.PanelOpen = False
        cpOpciones.CanExpand = False
    Else
        ' Ensayo de Tiempo
        cpOpciones.PanelOpen = True
        cpOpciones.CanExpand = True
        If lista.ListItems(lista.selectedItem.Index).SubItems(7) = "" Then
            ddesde = "01-01-1900"
        Else
            ddesde = Format(lista.ListItems(lista.selectedItem.Index).SubItems(7), "dd-mm-yyyy")
        End If
        If lista.ListItems(lista.selectedItem.Index).SubItems(8) = "" Then
            dhdesde.Value = Date & " 00:00:00"
        Else
            dhdesde.Value = Date & " " & lista.ListItems(lista.selectedItem.Index).SubItems(8)
        End If
        If lista.ListItems(lista.selectedItem.Index).SubItems(9) = "" Then
            dhasta = "01-01-1900"
        Else
            dhasta = Format(lista.ListItems(lista.selectedItem.Index).SubItems(9), "dd-mm-yyyy")
        End If
        If lista.ListItems(lista.selectedItem.Index).SubItems(10) = "" Then
            dhhasta.Value = Date & " 00:00:00"
        Else
            dhhasta.Value = Date & " " & lista.ListItems(lista.selectedItem.Index).SubItems(10)
        End If
    End If
    
    Dim oCe_resultados As New clsCe_resultados
    Dim rs As ADODB.Recordset
    Set rs = oCe_resultados.Listado_Pendientes_Muestra(CLng(lista.ListItems(lista.selectedItem.Index).SubItems(1)))
    Dim i As Integer
        probetas.ListItems.Clear
        If rs.RecordCount > 0 Then
            Do
                With probetas.ListItems.Add(, , rs(0)) ' id_muestra
                    .SubItems(1) = rs(1) ' designacion
                    .SubItems(2) = rs(2) ' probeta
                    .SubItems(3) = rs(3) ' area
                    .SubItems(4) = Trim(rs(5)) ' id_canagrosa
                    If rs(7) = "" Then
                        .SubItems(5) = " "
                    Else
                        .SubItems(5) = Format(rs(7), "dd-mm-yyyy") ' Fecha
                    End If
                    If rs(8) <> "" Then
                        .SubItems(6) = rs(8)
                    Else
                        .SubItems(6) = " "
                    End If
                    .SubItems(7) = rs(9)
                End With
                '
'                probetas.ListItems(probetas.ListItems.Count).SmallIcon = 0
'                If Trim(rs(7)) <> "" Then ' Fecha
'                    If rs(9) = 1 Then ' Resultado
'                        probetas.ListItems(probetas.ListItems.Count).SmallIcon = 2
'                    Else
'                        probetas.ListItems(probetas.ListItems.Count).SmallIcon = 1
'                    End If
'                Else
'                    probetas.ListItems(probetas.ListItems.Count).SmallIcon = 3
'                End If
                
                rs.MoveNext
            Loop Until rs.EOF
            probetas_Click
         End If
    
    Set rs = Nothing
End Sub

Private Sub lista_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
   If lista.ListItems.Count > 0 Then
     lista.SortKey = ColumnHeader.Index - 1
     If lista.SortOrder = 0 Then
        lista.SortOrder = 1
     Else
        lista.SortOrder = 0
     End If
     lista.Sorted = True
   End If
End Sub
Private Sub lista_KeyUp(KeyCode As Integer, Shift As Integer)
    lista_Click
End Sub

Private Sub probetas_Click()
    If lista.ListItems.Count > 0 And probetas.ListItems.Count > 0 Then
        ' Conforme
        If CInt(lista.ListItems(lista.selectedItem.Index).SubItems(2)) = 0 Then
            txtcriterio = lista.ListItems(lista.selectedItem.Index).SubItems(3)
            txtRangoMin = lista.ListItems(lista.selectedItem.Index).SubItems(4)
            txtRangoMax = lista.ListItems(lista.selectedItem.Index).SubItems(5)
            txtconformeresultado = probetas.ListItems(probetas.selectedItem.Index).SubItems(6)
            If Trim(probetas.ListItems(probetas.selectedItem.Index).SubItems(5)) <> "" Then
                fechaconforme.Value = probetas.ListItems(probetas.selectedItem.Index).SubItems(5)
                If probetas.ListItems(probetas.selectedItem.Index).SubItems(7) = 1 Then
                    chkConforme(0).Value = True
'                    probetas.ListItems(probetas.SelectedItem.Index).SmallIcon = 2
                Else
'                    probetas.ListItems(probetas.SelectedItem.Index).SmallIcon = 1
                    chkConforme(1).Value = False
                End If
            Else
                ' No hay resultado
                chkConforme(0).Value = False
                chkConforme(1).Value = False
                fechaconforme.Value = Date
'                probetas.ListItems(probetas.SelectedItem.Index).SmallIcon = 3
            End If
        Else
        ' Formula
            cargar_campos
        End If
    End If
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
                .setMUESTRA_ID = lista.ListItems(lista.selectedItem.Index).SubItems(1)
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
                If Trim(auxdatos.ListItems(i).SubItems(3)) <> "" Then
                    .Insertar
                End If
            End With
        End If
    Next
    ' Almacena en CE_resultados la Solucion
    Dim oCe_resultados As New clsCe_resultados
    With oCe_resultados
        For i = 1 To auxdatos.ListItems.Count
         If UCase(lblestado.Caption) = "DUPLICADA" Then
            If auxdatos.ListItems(i).SubItems(6) = "M" Then
              If Trim(auxdatos.ListItems(i).SubItems(3)) <> "" Then
                .setRESULTADO = Replace(auxdatos.ListItems(i).SubItems(3), ",", ".")
                .setFECHA = Format(Date, "yyyy-mm-dd")
                .Modificar_Resultado lista.ListItems(lista.selectedItem.Index).SubItems(1), auxdatos.ListItems(i).Text, auxdatos.ListItems(i).SubItems(1), auxdatos.ListItems(i).SubItems(2), False
              End If
            End If
         Else
            If auxdatos.ListItems(i).bold = True Then
              If Trim(auxdatos.ListItems(i).SubItems(3)) <> "" Then
                .setRESULTADO = Replace(auxdatos.ListItems(i).SubItems(3), ",", ".")
                .setFECHA = Format(Date, "yyyy-mm-dd")
                .Modificar_Resultado lista.ListItems(lista.selectedItem.Index).SubItems(1), auxdatos.ListItems(i).Text, auxdatos.ListItems(i).SubItems(1), auxdatos.ListItems(i).SubItems(2), False
              End If
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
Private Sub pasar_siguiente_campo()
    If datos.ListItems.Count > datos.selectedItem.Index Then
        Set datos.selectedItem = datos.ListItems(datos.selectedItem.Index + 1)
        datos_Click
    Else
        If probetas.ListItems.Count > probetas.selectedItem.Index Then
            Set probetas.selectedItem = probetas.ListItems(probetas.selectedItem.Index + 1)
            probetas_Click
            datos_Click
        Else
            txtdato = ""
            txtValor = ""
            datos.SetFocus
        End If
    End If
End Sub
Private Sub cargar_campos()
    Dim ocampos As New clsFormulas_campos
    Dim rs As New ADODB.Recordset
    Dim consulta As String
    Dim duplicado As Integer
    Dim nombre As String
    Dim i As Integer
    Dim j As Integer
    datos.ListItems.Clear
    cmdCalcular.Enabled = False
    Set rs = ocampos.ListaFormulas(lista.ListItems(lista.selectedItem.Index).SubItems(2)) ' ID_FORMULA
    lblestado.Caption = ""
    lblestado.visible = False
    If chkDuplicada.Value = Checked Then
        duplicado = 2
        lblestado.Caption = "DUPLICADA"
        lblestado.visible = True
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
                    If oCE_RD.Carga(lista.ListItems(lista.selectedItem.Index).SubItems(1), probetas.ListItems(probetas.selectedItem.Index).SubItems(1), probetas.ListItems(probetas.selectedItem.Index).SubItems(2), probetas.ListItems(probetas.selectedItem.Index).SubItems(3), rs("id_campo")) Then
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
     End If
     visualizar_duplicados
    End If
    ' Comprobar si ya tiene datos
    For i = 1 To auxdatos.ListItems.Count
        If probetas.ListItems(probetas.selectedItem.Index).SubItems(1) = auxdatos.ListItems(i) And _
           probetas.ListItems(probetas.selectedItem.Index).SubItems(2) = auxdatos.ListItems(i).SubItems(1) And _
           probetas.ListItems(probetas.selectedItem.Index).SubItems(3) = auxdatos.ListItems(i).SubItems(2) Then
            datos.ListItems(CInt(auxdatos.ListItems(i).SubItems(4))).SubItems(1) = auxdatos.ListItems(i).SubItems(3)
        End If
    Next
    Set rs = Nothing
    Set rs_campos = Nothing
    Set ocampos = Nothing
    datos_Click
End Sub
Private Sub grabar_auxdatos()
    Dim i As Integer
    For i = auxdatos.ListItems.Count To 1 Step -1
       If probetas.ListItems(probetas.selectedItem.Index).SubItems(1) = auxdatos.ListItems(i) And _
          probetas.ListItems(probetas.selectedItem.Index).SubItems(2) = auxdatos.ListItems(i).SubItems(1) And _
          probetas.ListItems(probetas.selectedItem.Index).SubItems(3) = auxdatos.ListItems(i).SubItems(2) Then
           auxdatos.ListItems.Remove (i)
       End If
    Next
    For i = 1 To datos.ListItems.Count
       With auxdatos.ListItems.Add(, , probetas.ListItems(probetas.selectedItem.Index).SubItems(1)) ' DESIGNACION
             .SubItems(1) = probetas.ListItems(probetas.selectedItem.Index).SubItems(2) ' PROBETA
             .SubItems(2) = probetas.ListItems(probetas.selectedItem.Index).SubItems(3) ' AREA
             .SubItems(3) = datos.ListItems(i).SubItems(1) ' VALOR
             .SubItems(4) = i ' LINEA
             .SubItems(5) = datos.ListItems(i).SubItems(3) ' CAMPO
             If datos.ListItems(i).bold = True Then
                .bold = True
                ' Si es solucion, la subimoslas determinaciones
                If UCase(lblestado.Caption) <> "DUPLICADA" Then
                    If datos.ListItems(i).SubItems(1) <> "" Then
                        probetas.ListItems(probetas.selectedItem.Index).SubItems(5) = Format(Date, "dd-mm-yyyy")
                        probetas.ListItems(probetas.selectedItem.Index).SubItems(6) = datos.ListItems(i).SubItems(1)
                    End If
                End If
             Else
                If UCase(lblestado.Caption) = "DUPLICADA" Then
                    If datos.ListItems(i).Text = "Resultado (MEDIA)" Then
                        .SubItems(6) = "M"
                    End If
                    If datos.ListItems(datos.ListItems.Count - 1).SubItems(1) <> "" Then
                        probetas.ListItems(probetas.selectedItem.Index).SubItems(5) = Format(Date, "dd-mm-yyyy")
                        probetas.ListItems(probetas.selectedItem.Index).SubItems(6) = datos.ListItems(datos.ListItems.Count - 1).SubItems(1)
                    End If
                End If
             End If
       End With
    Next
    almacenar_resultados_determinaciones
    Dim oMuestra As New clsMuestra
    oMuestra.comprobar_cierre CLng(lista.ListItems(lista.selectedItem.Index).SubItems(1))
    Set oMuestra = Nothing
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
        If numero_resultados = 2 And IsNumeric(res1) And IsNumeric(res2) Then ' Calcular media y diferencia
            Dim media As Single
            Dim dif As Single
            media = (CSng(res1) + CSng(res2)) / 2
            datos.ListItems(datos.ListItems.Count - 1).SubItems(1) = Format(CStr(media), "##0.00")
            grabar_auxdatos
            dif = Abs((CSng(res1) - CSng(res2)))
            datos.ListItems(datos.ListItems.Count).SubItems(1) = Format(CStr(dif), "#,##0.00")
            grabar_auxdatos
        Else
            If res1 = "--" Or res2 = "--" Then
                datos.ListItems(datos.ListItems.Count - 1).SubItems(1) = "--"
                datos.ListItems(datos.ListItems.Count).SubItems(1) = "--"
            Else
                If UCase(lblestado.Caption) = "DUPLICADA" Then
                    datos.ListItems(datos.ListItems.Count - 1).SubItems(1) = probetas.ListItems(probetas.selectedItem.Index).SubItems(6)
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
        txtValor = Trim(datos.ListItems(datos.selectedItem.Index).SubItems(1))
        txtValor.SetFocus
        txtValor.SelStart = 0
        txtValor.SelLength = Len(txtValor)
        txtdato = datos.ListItems(datos.selectedItem.Index)
    End If
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
'    Dim oDeter As New clsDeterminaciones
'    Dim oTD As New clsTipos_determinacion
'    oDeter.CargarDeterminacion (lista.ListItems(lista.SelectedItem.Index).Text)
'    oTD.CargarTipoDeterminacion (oDeter.getTIPO_DETERMINACION_ID)
    ofor.CARGAR (lista.ListItems(lista.selectedItem.Index).SubItems(2))
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
    datos.ListItems(datos.selectedItem.Index).SubItems(1) = formatear(sc.Eval(Formula), datos.ListItems(datos.selectedItem.Index).SubItems(4), datos.ListItems(datos.selectedItem.Index).SubItems(5))
    grabar_auxdatos
    visualizar_duplicados
    pasar_siguiente_campo
    Exit Sub
fallo:
    MsgBox "Error en la formula. " & Err.Description, vbCritical, "Error"

End Sub

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
