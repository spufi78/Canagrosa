VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{CEA3B7F6-2847-4E5E-A551-DB7A62489D44}#46.0#0"; "miCombo.ocx"
Begin VB.Form frmREX_Gestion 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Listado de Botes de Reactivos Externos"
   ClientHeight    =   9555
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   15330
   Icon            =   "frmREX_Gestion.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   9555
   ScaleWidth      =   15330
   Begin VB.CommandButton cmdAdjuntar 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Adjuntar"
      Height          =   870
      Left            =   12420
      Style           =   1  'Graphical
      TabIndex        =   75
      ToolTipText     =   "Genera una copia de la muestra seleccionada"
      Top             =   8640
      Width           =   1095
   End
   Begin VB.CommandButton cmdVerPedido 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Ver Pedido"
      Height          =   870
      Left            =   8280
      Picture         =   "frmREX_Gestion.frx":08CA
      Style           =   1  'Graphical
      TabIndex        =   65
      Top             =   8640
      Width           =   1005
   End
   Begin VB.CommandButton cmdNoConforme 
      BackColor       =   &H00E0E0E0&
      Caption         =   "No Conforme"
      Height          =   870
      Left            =   7245
      Picture         =   "frmREX_Gestion.frx":1194
      Style           =   1  'Graphical
      TabIndex        =   60
      Top             =   8640
      Width           =   1005
   End
   Begin VB.CommandButton cmdInventario 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Informes"
      Height          =   870
      Left            =   9315
      Picture         =   "frmREX_Gestion.frx":1A5E
      Style           =   1  'Graphical
      TabIndex        =   57
      Top             =   8640
      Width           =   1005
   End
   Begin VB.TextBox txtfiltro 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Height          =   330
      Index           =   2
      Left            =   6390
      TabIndex        =   6
      Top             =   1395
      Width           =   1770
   End
   Begin VB.CommandButton cmdMarcar 
      BackColor       =   &H00C0C0FF&
      Caption         =   "Marcar Todos"
      Height          =   330
      Left            =   45
      Style           =   1  'Graphical
      TabIndex        =   47
      Top             =   2430
      Width           =   1455
   End
   Begin VB.CommandButton cmdDesmarcar 
      BackColor       =   &H00C0C0FF&
      Caption         =   "Desmarcar Todos"
      Height          =   330
      Left            =   1485
      Style           =   1  'Graphical
      TabIndex        =   46
      Top             =   2430
      Width           =   1455
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Evaluación Certificado"
      Height          =   870
      Left            =   6210
      Picture         =   "frmREX_Gestion.frx":2328
      Style           =   1  'Graphical
      TabIndex        =   45
      Top             =   8640
      Width           =   1005
   End
   Begin VB.CommandButton cmdCertificadoExterno 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Certificado Externo"
      Height          =   870
      Left            =   5175
      Picture         =   "frmREX_Gestion.frx":2BF2
      Style           =   1  'Graphical
      TabIndex        =   44
      Top             =   8640
      Width           =   1005
   End
   Begin VB.CommandButton cmdExistencias 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Listado Existencias"
      Height          =   870
      Left            =   10350
      Picture         =   "frmREX_Gestion.frx":34BC
      Style           =   1  'Graphical
      TabIndex        =   36
      Top             =   8640
      Width           =   1005
   End
   Begin VB.CommandButton cmdPanreac 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Cert. Panreac"
      Height          =   870
      Left            =   4170
      Picture         =   "frmREX_Gestion.frx":3D86
      Style           =   1  'Graphical
      TabIndex        =   35
      Top             =   8640
      Width           =   1005
   End
   Begin VB.CommandButton cmdImprimir 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Imprimir"
      Height          =   870
      Left            =   11385
      Style           =   1  'Graphical
      TabIndex        =   34
      Top             =   8640
      Width           =   1005
   End
   Begin VB.CommandButton cmdSalir 
      BackColor       =   &H00E0E0E0&
      Caption         =   "ESC-Salir"
      Height          =   870
      Left            =   14220
      Style           =   1  'Graphical
      TabIndex        =   33
      Top             =   8640
      Width           =   1050
   End
   Begin VB.CommandButton cmdModificar 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Modificar Datos"
      Enabled         =   0   'False
      Height          =   870
      Left            =   3150
      Style           =   1  'Graphical
      TabIndex        =   32
      Top             =   8640
      Width           =   1005
   End
   Begin VB.CommandButton cmdmanual 
      Caption         =   "Código Manual"
      Height          =   585
      Left            =   9090
      TabIndex        =   31
      Top             =   8145
      Visible         =   0   'False
      Width           =   645
   End
   Begin VB.CommandButton cmdetiqueta 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Etiqueta"
      Enabled         =   0   'False
      Height          =   870
      Left            =   2130
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   8640
      Width           =   1005
   End
   Begin VB.CommandButton cmdTerminar 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Terminar"
      Enabled         =   0   'False
      Height          =   870
      Left            =   1110
      Picture         =   "frmREX_Gestion.frx":3FAB
      Style           =   1  'Graphical
      TabIndex        =   18
      Top             =   8640
      Width           =   1005
   End
   Begin VB.CommandButton cmdAbrir 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Abrir"
      Enabled         =   0   'False
      Height          =   870
      Left            =   90
      Picture         =   "frmREX_Gestion.frx":4E75
      Style           =   1  'Graphical
      TabIndex        =   17
      Top             =   8640
      Width           =   1005
   End
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
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
      ForeColor       =   &H80000008&
      Height          =   2070
      Left            =   45
      TabIndex        =   19
      Top             =   360
      Width           =   15240
      Begin VB.Frame Frame2 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Caption         =   "HENKEL"
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
         Height          =   555
         Index           =   3
         Left            =   9630
         TabIndex        =   71
         Top             =   1425
         Width           =   1980
         Begin VB.OptionButton opHenkel 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            Caption         =   "Si"
            ForeColor       =   &H80000008&
            Height          =   285
            Index           =   1
            Left            =   900
            TabIndex        =   73
            Top             =   225
            Width           =   465
         End
         Begin VB.OptionButton opHenkel 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            Caption         =   "Todos"
            ForeColor       =   &H80000008&
            Height          =   285
            Index           =   0
            Left            =   90
            TabIndex        =   74
            Top             =   225
            Value           =   -1  'True
            Width           =   780
         End
         Begin VB.OptionButton opHenkel 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            Caption         =   "No"
            ForeColor       =   &H80000008&
            Height          =   285
            Index           =   2
            Left            =   1350
            TabIndex        =   72
            Top             =   225
            Width           =   555
         End
      End
      Begin VB.CheckBox chkProbetas 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Caption         =   "Probetas"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   10530
         TabIndex        =   66
         Top             =   405
         Width           =   1185
      End
      Begin VB.Frame Frame2 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Caption         =   "C.A.U."
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
         Height          =   555
         Index           =   2
         Left            =   7695
         TabIndex        =   61
         Top             =   1425
         Width           =   1890
         Begin VB.OptionButton opNoConforme 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            Caption         =   "No"
            ForeColor       =   &H80000008&
            Height          =   285
            Index           =   0
            Left            =   1305
            TabIndex        =   64
            Top             =   225
            Width           =   510
         End
         Begin VB.OptionButton opNoConforme 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            Caption         =   "Si"
            ForeColor       =   &H80000008&
            Height          =   285
            Index           =   1
            Left            =   855
            TabIndex        =   63
            Top             =   225
            Width           =   465
         End
         Begin VB.OptionButton opNoConforme 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            Caption         =   "Todos"
            ForeColor       =   &H80000008&
            Height          =   285
            Index           =   2
            Left            =   90
            TabIndex        =   62
            Top             =   225
            Value           =   -1  'True
            Width           =   825
         End
      End
      Begin VB.Frame Frame6 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Caption         =   "Código"
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
         Height          =   600
         Left            =   13545
         TabIndex        =   58
         Top             =   1395
         Width           =   1575
         Begin VB.TextBox txtcodigo 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            Height          =   300
            Left            =   90
            TabIndex        =   59
            Top             =   225
            Width           =   1410
         End
      End
      Begin VB.CheckBox chkfechas 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Check1"
         Height          =   285
         Left            =   90
         TabIndex        =   52
         Top             =   675
         Width           =   195
      End
      Begin VB.TextBox txtfiltro 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   330
         Index           =   1
         Left            =   3915
         TabIndex        =   5
         Top             =   1035
         Width           =   1320
      End
      Begin VB.TextBox txtfiltro 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   330
         Index           =   0
         Left            =   1350
         TabIndex        =   4
         Top             =   1035
         Width           =   1320
      End
      Begin VB.Frame Frame2 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Caption         =   "Tipo de Reactivo"
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
         Height          =   1995
         Index           =   1
         Left            =   11790
         TabIndex        =   37
         Top             =   0
         Width           =   1740
         Begin VB.CheckBox chktiporeactivo 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            Caption         =   "Prod.Controlado"
            ForeColor       =   &H80000008&
            Height          =   195
            Index           =   6
            Left            =   135
            TabIndex        =   53
            Top             =   1665
            Width           =   1545
         End
         Begin VB.CheckBox chktiporeactivo 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            Caption         =   "R.C."
            ForeColor       =   &H80000008&
            Height          =   195
            Index           =   5
            Left            =   135
            TabIndex        =   43
            Top             =   1440
            Width           =   1545
         End
         Begin VB.CheckBox chktiporeactivo 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            Caption         =   "Otros"
            ForeColor       =   &H80000008&
            Height          =   195
            Index           =   4
            Left            =   135
            TabIndex        =   42
            Top             =   1215
            Width           =   1545
         End
         Begin VB.CheckBox chktiporeactivo 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            Caption         =   "Mat. Fungible"
            ForeColor       =   &H80000008&
            Height          =   195
            Index           =   3
            Left            =   135
            TabIndex        =   41
            Top             =   990
            Width           =   1545
         End
         Begin VB.CheckBox chktiporeactivo 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            Caption         =   "M.R.C."
            ForeColor       =   &H80000008&
            Height          =   195
            Index           =   2
            Left            =   135
            TabIndex        =   40
            Top             =   765
            Width           =   1545
         End
         Begin VB.CheckBox chktiporeactivo 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            Caption         =   "M.R."
            ForeColor       =   &H80000008&
            Height          =   195
            Index           =   1
            Left            =   135
            TabIndex        =   39
            Top             =   540
            Width           =   1545
         End
         Begin VB.CheckBox chktiporeactivo 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            Caption         =   "Reactivo Normal"
            ForeColor       =   &H80000008&
            Height          =   195
            Index           =   0
            Left            =   135
            TabIndex        =   38
            Top             =   315
            Value           =   1  'Checked
            Width           =   1545
         End
      End
      Begin VB.Frame Frame5 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Caption         =   "Caducados"
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
         Height          =   555
         Left            =   3915
         TabIndex        =   28
         Top             =   1425
         Width           =   1920
         Begin VB.OptionButton opCaducado 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            Caption         =   "Todos"
            ForeColor       =   &H80000008&
            Height          =   285
            Index           =   2
            Left            =   90
            TabIndex        =   54
            Top             =   225
            Width           =   780
         End
         Begin VB.OptionButton opCaducado 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            Caption         =   "No"
            ForeColor       =   &H80000008&
            Height          =   285
            Index           =   1
            Left            =   1350
            TabIndex        =   30
            Top             =   225
            Value           =   -1  'True
            Width           =   525
         End
         Begin VB.OptionButton opCaducado 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            Caption         =   "Si"
            ForeColor       =   &H80000008&
            Height          =   285
            Index           =   0
            Left            =   870
            TabIndex        =   29
            Top             =   225
            Width           =   480
         End
      End
      Begin VB.Frame Frame3 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Caption         =   "Abierto"
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
         Height          =   555
         Left            =   60
         TabIndex        =   26
         Top             =   1425
         Width           =   1875
         Begin VB.OptionButton opAbierto 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            Caption         =   "No"
            ForeColor       =   &H80000008&
            Height          =   285
            Index           =   2
            Left            =   1305
            TabIndex        =   9
            Top             =   225
            Width           =   525
         End
         Begin VB.OptionButton opAbierto 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            Caption         =   "Si"
            ForeColor       =   &H80000008&
            Height          =   285
            Index           =   1
            Left            =   810
            TabIndex        =   8
            Top             =   225
            Width           =   510
         End
         Begin VB.OptionButton opAbierto 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            Caption         =   "Todos"
            ForeColor       =   &H80000008&
            Height          =   285
            Index           =   0
            Left            =   45
            TabIndex        =   7
            Top             =   225
            Value           =   -1  'True
            Width           =   780
         End
      End
      Begin VB.CheckBox chkTodosReactivos 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Caption         =   "Todos"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   10530
         TabIndex        =   1
         Top             =   135
         Value           =   1  'Checked
         Width           =   960
      End
      Begin VB.Frame Frame2 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Caption         =   "Anulados"
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
         Height          =   555
         Index           =   0
         Left            =   5850
         TabIndex        =   20
         Top             =   1425
         Width           =   1830
         Begin VB.OptionButton opTipo 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            Caption         =   "Si"
            ForeColor       =   &H80000008&
            Height          =   285
            Index           =   1
            Left            =   810
            TabIndex        =   14
            Top             =   225
            Width           =   465
         End
         Begin VB.OptionButton opTipo 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            Caption         =   "Todos"
            ForeColor       =   &H80000008&
            Height          =   285
            Index           =   2
            Left            =   45
            TabIndex        =   55
            Top             =   225
            Width           =   825
         End
         Begin VB.OptionButton opTipo 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            Caption         =   "No"
            ForeColor       =   &H80000008&
            Height          =   285
            Index           =   0
            Left            =   1260
            TabIndex        =   13
            Top             =   225
            Value           =   -1  'True
            Width           =   510
         End
      End
      Begin MSComCtl2.DTPicker fdesde 
         Height          =   330
         Left            =   1350
         TabIndex        =   2
         Top             =   630
         Width           =   1320
         _ExtentX        =   2328
         _ExtentY        =   582
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
         Format          =   51970049
         CurrentDate     =   38002
      End
      Begin MSComCtl2.DTPicker fhasta 
         Height          =   330
         Left            =   3915
         TabIndex        =   3
         Top             =   630
         Width           =   1320
         _ExtentX        =   2328
         _ExtentY        =   582
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
         Format          =   51970049
         CurrentDate     =   38002
      End
      Begin VB.Frame Frame4 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Caption         =   "Terminado"
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
         Height          =   555
         Left            =   1980
         TabIndex        =   27
         Top             =   1425
         Width           =   1920
         Begin VB.OptionButton opTerminado 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            Caption         =   "No"
            ForeColor       =   &H80000008&
            Height          =   285
            Index           =   2
            Left            =   1305
            TabIndex        =   12
            Top             =   225
            Value           =   -1  'True
            Width           =   525
         End
         Begin VB.OptionButton opTerminado 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            Caption         =   "Si"
            ForeColor       =   &H80000008&
            Height          =   285
            Index           =   1
            Left            =   825
            TabIndex        =   11
            Top             =   225
            Width           =   555
         End
         Begin VB.OptionButton opTerminado 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            Caption         =   "Todos"
            ForeColor       =   &H80000008&
            Height          =   285
            Index           =   0
            Left            =   90
            TabIndex        =   10
            Top             =   225
            Width           =   825
         End
      End
      Begin pryCombo.miCombo cmbReactivos 
         Height          =   330
         Left            =   1350
         TabIndex        =   56
         Top             =   270
         Width           =   8925
         _ExtentX        =   15743
         _ExtentY        =   582
      End
      Begin MSDataListLib.DataCombo cmbCentro 
         Bindings        =   "frmREX_Gestion.frx":5D3F
         Height          =   315
         Left            =   8865
         TabIndex        =   67
         Top             =   1035
         Width           =   1800
         _ExtentX        =   3175
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
      Begin pryCombo.miCombo cmbProveedores 
         Height          =   330
         Left            =   6345
         TabIndex        =   69
         Top             =   675
         Width           =   5370
         _ExtentX        =   9472
         _ExtentY        =   582
      End
      Begin VB.CommandButton cmdBuscar 
         BackColor       =   &H00E0E0E0&
         Caption         =   "&Buscar"
         Height          =   915
         Left            =   13860
         Style           =   1  'Graphical
         TabIndex        =   15
         Top             =   315
         Width           =   1050
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Proveedor"
         Height          =   195
         Index           =   0
         Left            =   5490
         TabIndex        =   70
         Top             =   720
         Width           =   780
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Centro"
         Height          =   195
         Index           =   22
         Left            =   8190
         TabIndex        =   68
         Top             =   1080
         Width           =   465
      End
      Begin VB.Label lblCampos 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Lote"
         Height          =   195
         Index           =   12
         Left            =   5490
         TabIndex        =   51
         Top             =   1080
         Width           =   345
      End
      Begin VB.Label lblCampos 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Reactivo"
         Height          =   240
         Index           =   10
         Left            =   3060
         TabIndex        =   49
         Top             =   1080
         Width           =   675
      End
      Begin VB.Label lblCampos 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Código Bote"
         Height          =   195
         Index           =   9
         Left            =   90
         TabIndex        =   48
         Top             =   1080
         Width           =   885
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Reactivo"
         Height          =   195
         Index           =   3
         Left            =   135
         TabIndex        =   25
         Top             =   360
         Width           =   645
      End
      Begin VB.Label lblCampos 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "hasta"
         Height          =   195
         Index           =   2
         Left            =   3060
         TabIndex        =   22
         Top             =   720
         Width           =   450
      End
      Begin VB.Label lblCampos 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Recep. desde"
         Height          =   195
         Index           =   1
         Left            =   270
         TabIndex        =   21
         Top             =   720
         Width           =   1035
      End
   End
   Begin MSComctlLib.ListView lista 
      Height          =   5865
      Left            =   45
      TabIndex        =   0
      Top             =   2745
      Width           =   15270
      _ExtentX        =   26935
      _ExtentY        =   10345
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      AllowReorder    =   -1  'True
      Checkboxes      =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      Icons           =   "ImageList1"
      SmallIcons      =   "ImageList1"
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
      Left            =   13500
      Top             =   8820
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
            Picture         =   "frmREX_Gestion.frx":5D85
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmREX_Gestion.frx":621B
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmREX_Gestion.frx":66B2
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Label lblCampos 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      Caption         =   "Código"
      Height          =   195
      Index           =   11
      Left            =   6075
      TabIndex        =   50
      Top             =   1755
      Width           =   525
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Caption         =   "Listado de Botes de Reactivos Externos"
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
      Index           =   4
      Left            =   45
      TabIndex        =   24
      Top             =   0
      Width           =   15240
   End
   Begin VB.Label lblmsg 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
      Caption         =   "Seleccione Criterio."
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   45
      TabIndex        =   23
      Top             =   2430
      Width           =   15240
   End
End
Attribute VB_Name = "frmREX_Gestion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" _
        (ByVal Hwnd As Long, ByVal lpOperation As String, _
        ByVal lpFile As String, ByVal lpParameters As String, _
        ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Private mvarstrCriterioInforme As String


Private Sub cmdAdjuntar_Click()
    If lista.ListItems.Count = 0 Then Exit Sub
    Dim m As String
    Dim i As Integer
    For i = 1 To lista.ListItems.Count
        If lista.ListItems(i).Checked = True Then
            m = m & lista.ListItems(i).Text & ";"
       End If
    Next
    With frmAdjuntos
        .TOBJETO = TOBJETO.TOBJETO_REX_CERTIFICADOS
        If m = "" Then
            .COBJETO = lista.ListItems(lista.selectedItem.Index).Text
        Else
            .COBJETO = 0
        End If
        .COBJETO_GRUPO_MUESTRAS = m
        .Show 1
    End With
    Set frmAdjuntos = Nothing
End Sub


Private Sub cmdVerPedido_Click()
    'M1076-I
   On Error GoTo cmdVerPedido_Click_Error

    If lista.ListItems.Count > 0 Then
        Dim oBote As New clsBotes_ex
        If oBote.CARGAR(CLng(lista.ListItems(lista.selectedItem.Index).Text)) = True Then
            frmREX_Pedidos_Detalle.PK = oBote.getPEDIDO_BOTE_EX_ID
            frmREX_Pedidos_Detalle.TIPO_BOTE_ID = oBote.getTIPO_BOTE_EX_ID
            frmREX_Pedidos_Detalle.Show 1
        End If
    End If
    'M1076-F

   On Error GoTo 0
   Exit Sub

cmdVerPedido_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure cmdVerPedido_Click of Formulario frmREX_Gestion"
End Sub

Private Sub cmbReactivos_change()
    If cmbReactivos.getTEXTO <> "" Then
        cmdBuscar_Click
    End If
End Sub

Private Sub cmdImprimir_Click()
    Dim oBREX As New clsBotes_ex
    oBREX.Imprimir_Listado mvarstrCriterioInforme
    Set oBREX = Nothing
End Sub

Private Sub chkfechas_Click()
    If chkFechas.Value = Checked Then
        fdesde.Enabled = True
        fhasta.Enabled = True
    Else
        fdesde.Enabled = False
        fhasta.Enabled = False
    End If
End Sub

Private Sub chktiporeactivo_Click(Index As Integer)
    cargar_reactivos
    cmdBuscar_Click
End Sub

Private Sub cmdDesmarcar_Click()
    Dim i As Integer
    For i = 1 To lista.ListItems.Count
        lista.ListItems(i).Checked = False
    Next
End Sub

Private Sub cmdInventario_Click()
    frmREX_Informes.Show 1
End Sub

Private Sub cmdMarcar_Click()
    Dim i As Integer
    For i = 1 To lista.ListItems.Count
        lista.ListItems(i).Checked = True
    Next
End Sub

Private Sub chkTodosReactivos_Click()
    If chkTodosReactivos.Value = Checked Then
        cmbReactivos.limpiar
        cmbReactivos.desactivar
    Else
        cmbReactivos.activar
    End If
End Sub

Private Sub cmdAbrir_Click()
    If lista.ListItems.Count > 0 Then
        ' Verificar si hay algun no conforme
        Dim i As Integer
        For i = 1 To lista.ListItems.Count
            If lista.ListItems(i).Checked = True Then
                If lista.ListItems(i).SmallIcon = 3 Then
                    MsgBox "La evaluación del reactivo, número : " & lista.ListItems(i).Text & " es NO CONFORME, no se puede abrir.", vbCritical, App.Title
                    Exit Sub
                End If
            End If
        Next
        
        Dim fecha As String
        fecha = InputBox("Introduzca la fecha de apertura para los botes marcados.", "Fecha apertura", Format(Date, "dd/mm/yyyy"))
        If fecha <> "" Then
            If IsDate(fecha) = False Then
                MsgBox "El formato de la fecha no es correcto.", vbCritical, App.Title
                Exit Sub
            End If
'        If MsgBox("¿Abrir los botes marcados con fecha de hoy?", vbQuestion + vbYesNo, App.Title) = vbYes Then
            Dim obe As New clsBotes_ex
            Dim se As Boolean
            se = False
            For i = 1 To lista.ListItems.Count
                If lista.ListItems(i).Checked = True Then
                    se = True
                    obe.Abrir lista.ListItems(i).Text, fecha
                End If
            Next
            If se = True Then
                If txtCodigo <> "" Then
                    txtcodigo_LostFocus
                Else
                    Call buscar
                End If
            Else
                MsgBox "No hay ningún bote marcado.", vbInformation, App.Title
            End If
        End If
    End If
    lista_Click
End Sub

Private Sub cmdCertificadoExterno_Click()
    On Error GoTo fallo
    If lista.ListItems.Count > 0 Then
' M0601-I
        Dim oAdjunto As New clsAdjuntos
        If oAdjunto.CargarDocumentoUltimo(TOBJETO.TOBJETO_REX_CERTIFICADOS, CLng(lista.ListItems(lista.selectedItem.Index).Text), 0, True, ADJUNTOS_TIPOS.ADJUNTOS_TIPOS_CERTIFICADO) = "" Then
            MsgBox "El certificado no esta informado. Informelo previamente en la opción Modificar Datos.", vbInformation, App.Title
        End If
        Set oAdjunto = Nothing
'        Dim oBote As New clsBotes_ex
'        oBote.CARGAR CLng(lista.ListItems(lista.selectedItem.Index).Text)
'        If oBote.getCERTIFICADO_EXTERNO <> "" Then
'            If Dir(oBote.getCERTIFICADO_EXTERNO) <> "" Then
'                r = Shell("rundll32.exe url.dll,FileProtocolHandler " & oBote.getCERTIFICADO_EXTERNO, vbMaximizedFocus)
'            Else
'                MsgBox "No se localiza el certificado. Informelo nuevamente en la opción Modificar Datos.", vbInformation, App.Title
'            End If
'        Else
'            MsgBox "El certificado no esta informado. Informelo previamente en la opción Modificar Datos.", vbInformation, App.Title
'        End If
'M0601-F
    End If
    Exit Sub
fallo:
    MsgBox "Error al abrir el documento.", vbCritical, App.Title

End Sub

Private Sub cmdetiqueta_Click()
    nueva_etiqueta
End Sub

Private Sub cmdExistencias_Click()
    opAbierto(0).Value = True
    opTerminado(2).Value = True
    opCaducado(2).Value = True
    opTipo(0).Value = True
    txtFiltro(0) = ""
    txtFiltro(1) = ""
    txtFiltro(2) = ""
    chkTodosReactivos.Value = Checked
    chkFechas.Value = Unchecked
    cmdBuscar_Click
End Sub

Private Sub cmdmanual_Click()
    If lista.ListItems.Count > 0 Then
        Dim consulta As String
        Dim NUEVO As String
        NUEVO = InputBox("Intro nuevo numero : ", App.Title)
        If NUEVO <> "" Then
            consulta = "update botes_ex set id_bote_ex = " & NUEVO & " where id_bote_ex = " & CLng(lista.ListItems(lista.selectedItem.Index).Text)
            execute_bd consulta
            cmdBuscar_Click
        End If
    End If
End Sub

Private Sub cmdModificar_Click()
'    gbotereactivoex = CLng(lista.ListItems(lista.SelectedItem.Index).Text)
    frmREX_Bote_Modificacion.PK = CLng(lista.ListItems(lista.selectedItem.Index).Text)
    frmREX_Bote_Modificacion.Show 1
    actualizar_lista
'    gbotereactivoex = 0
End Sub

Private Sub cmdNoConforme_Click()
    If lista.ListItems.Count > 0 Then
        If MsgBox("¿Desea poner como NO CONFORMES los botes marcados?", vbQuestion + vbYesNo, App.Title) = vbYes Then
            Dim i As Integer
            Dim obe As New clsBotes_ex
            Dim se As Boolean
            se = False
            For i = 1 To lista.ListItems.Count
                If lista.ListItems(i).Checked = True Then
                    se = True
                    obe.No_Conforme_Modificar lista.ListItems(i).Text
                End If
            Next
            If se = True Then
                If txtCodigo <> "" Then
                    txtcodigo_LostFocus
                Else
                    Call buscar
                End If
            Else
                MsgBox "No hay ningún bote marcado.", vbInformation, App.Title
            End If
        End If
    End If
    lista_Click
End Sub

Private Sub cmdPanreac_Click()
    If lista.ListItems.Count > 0 Then
        Dim oParametro As New clsParametros
        Dim oBote As New clsBotes_ex
        If oBote.CARGAR(lista.ListItems(lista.selectedItem.Index).Text) = True Then
            Dim oTipos_bote As New clsTipos_bote_ex
            If oTipos_bote.CARGAR(oBote.getTIPO_BOTE_EX_ID) = True Then
            oParametro.Carga parametros.PROVEEDOR_PANREAC, ""
            If CInt(oParametro.getVALOR) <> oTipos_bote.getPROVEEDOR_ID Then
                MsgBox "No es un producto de Panreac.", vbInformation, App.Title
                Exit Sub
            End If
            Dim iret As Long
            Dim cadena As String
            oParametro.Carga 1, ""
            cadena = oParametro.getVALOR
            cadena = cadena & Left(oTipos_bote.getCODIGO, 6) & "-" & oBote.getLOTE & "ES.htm"
            iret = ShellExecute(Me.Hwnd, vbNullString, cadena, vbNullString, "c:", SW_SHOWNORMAL)
            End If
        End If
    End If
End Sub

Private Sub cmdSalir_Click()
    Unload Me
End Sub

Private Sub cmdTerminar_Click()
    If lista.ListItems.Count > 0 Then
        Dim i As Integer
        For i = 1 To lista.ListItems.Count
            If lista.ListItems(i).Checked = True Then
                If lista.ListItems(i).SubItems(5) = "" Then
                    MsgBox "Existen botes que no tienen la fecha de apertura informada.", vbCritical, App.Title
                    Exit Sub
                End If
            End If
        Next
        Dim fecha As String
        fecha = InputBox("Introduzca la fecha de cierre para los botes marcados.", "Fecha cierre", Format(Date, "dd/mm/yyyy"))
        If fecha <> "" Then
            If IsDate(fecha) = False Then
                MsgBox "El formato de la fecha no es correcto.", vbCritical, App.Title
                Exit Sub
            End If
'        If MsgBox("¿Terminar los botes marcados con fecha de hoy?", vbQuestion + vbYesNo, App.Title) = vbYes Then
            Dim obe As New clsBotes_ex
            Dim se As Boolean
            se = False
            For i = 1 To lista.ListItems.Count
                If lista.ListItems(i).Checked = True Then
                    se = True
                    obe.Terminar lista.ListItems(i).Text, fecha
                End If
            Next
            If se = True Then
                If txtCodigo <> "" Then
                    txtcodigo_LostFocus
                Else
                    Call buscar
                End If
            Else
                MsgBox "No hay ningún bote marcado.", vbInformation, App.Title
            End If
        End If
    End If
    lista_Click
End Sub


Private Sub Command2_Click()
    If lista.ListItems.Count > 0 Then
          Dim oBote As New clsBotes_ex
          oBote.CARGAR lista.ListItems(lista.selectedItem.Index).Text
          If oBote.getCONVERTIDO = 0 Then
            frmREX_evaluacion.BOTE_EX_ID = lista.ListItems(lista.selectedItem.Index).Text
            frmREX_evaluacion.consulta = False
            frmREX_evaluacion.Show 1
          Else
            frmREX_Evaluacion_Parametros.BOTE_EX_ID = lista.ListItems(lista.selectedItem.Index).Text
            frmREX_Evaluacion_Parametros.Show 1
          End If
          actualizar_lista
    End If
End Sub

Private Sub fdesde_Change()
    cmdBuscar_Click
End Sub

Private Sub fhasta_Change()
    cmdBuscar_Click
End Sub

Private Sub Form_Activate()
    Me.SetFocus
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = 27 Then
        cmdSalir_Click
    End If
    If KeyCode = 112 Then
        cmdManual.visible = Not cmdManual.visible
    End If
End Sub
Private Sub Form_Load()
    log (Me.Name)
    cargar_botones Me
    Me.Left = 50
    Me.top = 50
    cabecera
    cmbReactivos.desactivar
    llenar_combo cmbProveedores, New clsProveedor, 0, frmProveedores_Detalle, ""
    cargar_combo cmbCentro, New clsCentros
    cargar_reactivos
'    llenar_combo cmbReactivos, New clsTipos_reactivo_ex, 0, frmREX_Reactivo, " ANULADO = 0 "
'    Dim oDeco As New clsDecodificadora
'    oDeco.cargar_combo cmbTipo, decodificadora.REX_TIPOS
    fdesde = Date
    fhasta = Date
    buscar
End Sub
Private Sub cabecera()
    With lista.ColumnHeaders.Add(, , "General", 1200, lvwColumnLeft)
        .Tag = "General"
    End With
    With lista.ColumnHeaders.Add(, , "Particular", 1150, lvwColumnCenter)
        .Tag = "Particular"
    End With
    With lista.ColumnHeaders.Add(, , "Cantidad", 1150, lvwColumnCenter)
        .Tag = "Cantidad"
    End With
    With lista.ColumnHeaders.Add(, , "Reactivo", 3850, lvwColumnLeft)
        .Tag = "Reactivo"
    End With
    With lista.ColumnHeaders.Add(, , "Recepción", 1050, lvwColumnCenter)
        .Tag = "Recepción"
    End With
    With lista.ColumnHeaders.Add(, , "Apertura", 1050, lvwColumnCenter)
        .Tag = "Apertura"
    End With
    With lista.ColumnHeaders.Add(, , "Terminado", 1050, lvwColumnCenter)
        .Tag = "Terminado"
    End With
    With lista.ColumnHeaders.Add(, , "Caducidad", 1050, lvwColumnCenter)
        .Tag = "Caducidad"
    End With
    With lista.ColumnHeaders.Add(, , "ID", 1, lvwColumnCenter)
        .Tag = "ID"
    End With
    With lista.ColumnHeaders.Add(, , "Lote", 2100, lvwColumnCenter)
        .Tag = "Lote"
    End With
    With lista.ColumnHeaders.Add(, , "Precio", 0, lvwColumnRight)
        .Tag = "Precio"
    End With
    With lista.ColumnHeaders.Add(, , "tipo_material_referencia", 1, lvwColumnRight)
        .Tag = "tipo_reactivo"
    End With
    With lista.ColumnHeaders.Add(, , "Centro", 1200, lvwColumnCenter)
        .Tag = "Centro"
    End With
End Sub
Private Sub cmdBuscar_Click()
    Call buscar
End Sub
Private Sub buscar()
    Dim consulta As String
    Dim strReactivo As String
    Dim strAbierto As String
'    Dim strCerrado As String
    Dim strAnulado As String
    Dim strCaducado As String
    Dim strconforme As String
    On Error GoTo fallo
    
    
    mvarstrCriterioInforme = ""
    
    lista.ListItems.Clear
    
    Dim rs As New ADODB.Recordset
    Dim f_desde As String
    Dim f_hasta As String
    f_desde = Format(fdesde, "yyyy-mm-dd")
    f_hasta = Format(fhasta, "yyyy-mm-dd")
    
    
    
    txtCodigo = ""
    ' Tipo de Bote
    strBote = ""
    ' Tipo reactivo
    strReactivo = ""
    If chkTodosReactivos.Value = Unchecked Then
        If cmbReactivos.getTEXTO = "" Then
'            MsgBox "Debe seleccionar un Reactivo.", vbExclamation, App.Title
            Exit Sub
        End If
        strReactivo = " AND tb.tipo_reactivo_ex_id = " & cmbReactivos.getPK_SALIDA
        mvarstrCriterioInforme = mvarstrCriterioInforme & " AND {tipos_bote_ex.tipo_reactivo_ex_id}=" & CStr(cmbReactivos.getPK_SALIDA)
    End If
    ' Fechas
    Dim fecha_desde As String
    Dim fecha_hasta As String
    If chkFechas.Value = Checked Then
        fecha_desde = " AND be.fecha_recepcion>='" & f_desde & "'"
        fecha_hasta = " AND be.fecha_recepcion<='" & f_hasta & "'"
        mvarstrCriterioInforme = mvarstrCriterioInforme & " AND {botes_ex.fecha_recepcion} >= date(" & Format(f_desde, "yyyy, mm, dd") & ") AND {botes_ex.fecha_recepcion} <= date(" & Format(f_hasta, "yyyy, mm, dd") & ")"
    End If
    
    'Abierto
    strAbierto = ""
    Dim aaux As String
    aaux = "0000-00-00"
    If opAbierto(1).Value = True Then
        strAbierto = " AND be.ABIERTO = 1"
        mvarstrCriterioInforme = mvarstrCriterioInforme & " AND {botes_ex.ABIERTO} = 1 "
    ElseIf opAbierto(2).Value = True Then
        strAbierto = " AND be.ABIERTO = 0"
        mvarstrCriterioInforme = mvarstrCriterioInforme & " AND {botes_ex.ABIERTO} = 0 "
    End If
    'Conforme
    strconforme = ""
    If opNoConforme(0).Value = True Then ' NO CONFORME
        strconforme = " AND be.no_conforme = 0"
        mvarstrCriterioInforme = mvarstrCriterioInforme & " AND {botes_ex.no_conforme} = 0"
    ElseIf opNoConforme(1).Value = True Then ' SI CONFORME
        strconforme = " AND be.no_conforme = 1"
        mvarstrCriterioInforme = mvarstrCriterioInforme & " AND {botes_ex.no_conforme} = 1"
    End If
    
    ' Terminado
    strTerminado = ""
    If opTerminado(1).Value = True Then
'        strTerminado = " AND be.fecha_fin <> 0"
'        mvarstrCriterioInforme = mvarstrCriterioInforme & " AND {botes_ex.fecha_fin} <> date(0000, 00, 00)"
        strTerminado = " AND be.finalizado = 1"
        mvarstrCriterioInforme = mvarstrCriterioInforme & " AND {botes_ex.finalizado} = 1"
    ElseIf opTerminado(2).Value = True Then
'        strTerminado = " AND be.fecha_fin = 0"
'        mvarstrCriterioInforme = mvarstrCriterioInforme & " AND {botes_ex.fecha_fin} = date(0000, 00, 00)"
        strTerminado = " AND be.finalizado = 0"
        mvarstrCriterioInforme = mvarstrCriterioInforme & " AND {botes_ex.finalizado} = 0"
    End If
    ' Caducado
    strCaducado = ""
    If opCaducado(0).Value = True Then ' Si caducados
        strCaducado = "                                     AND be.no_caduca = 0 AND be.fecha_caducidad < '" & Format(Date, "yyyy-mm-dd") & "'"
        mvarstrCriterioInforme = mvarstrCriterioInforme & " AND {botes_ex.no_caduca} = 0 AND {botes_ex.fecha_caducidad} < date(" & Format(Date, "yyyy, mm, dd") & ")"
    ElseIf opCaducado(1).Value = True Then
        strCaducado = "                                     AND (be.no_caduca = 1 or (be.no_caduca = 0 AND         be.fecha_caducidad > '" & Format(Date, "yyyy-mm-dd") & "'))"
        mvarstrCriterioInforme = mvarstrCriterioInforme & " AND ({botes_ex.no_caduca} = 1 or ({botes_ex.no_caduca} = 0 AND {botes_ex.fecha_caducidad} > date(" & Format(Date, "yyyy, mm, dd") & ")))"
    Else
        strCaducado = ""
    End If
    ' Anulado
    strAnulado = ""
    If opTipo(0).Value = True Then
        strAnulado = " AND be.anulado = 0"
        mvarstrCriterioInforme = mvarstrCriterioInforme & " AND {botes_ex.ANULADO} = 0"
    ElseIf opTipo(1).Value = True Then
        strAnulado = " AND be.anulado = 1"
        mvarstrCriterioInforme = mvarstrCriterioInforme & " AND {botes_ex.ANULADO} = 1"
    Else
        strAnulado = ""
    End If
    
    ' Mat. Referencia
    Dim matref As String
    Dim h As Integer
    Dim aux As String
    For h = 0 To 6
        If chktiporeactivo(h).Value = Checked Then
            aux = aux & h + 1 & ","
        End If
    Next
    If Len(aux) > 0 Then
        matref = " AND tb.tipo_m_referencia_id in (" & Left(aux, Len(aux) - 1) & ")"
        mvarstrCriterioInforme = mvarstrCriterioInforme & " AND {tipos_bote_ex.tipo_m_referencia_id} in [" & Left(aux, Len(aux) - 1) & "]"
    Else
        Exit Sub
    End If
    ' Filtros
    Dim filtro As String
    ' Codigo
    If Trim(Trim(txtFiltro(0))) <> "" Then
        filtro = filtro & " AND tb.codigo like '%" & Trim(txtFiltro(0)) & "%'"
        mvarstrCriterioInforme = mvarstrCriterioInforme & " AND {tipos_bote_ex.codigo} like " & Chr(34) & "*" & Trim(txtFiltro(0)) & "*" & Chr(34)
    End If
    ' Reactivo
    If Trim(Trim(txtFiltro(1))) <> "" Then
        filtro = filtro & " AND tr.nombre like '%" & Trim(txtFiltro(1)) & "%'"
        mvarstrCriterioInforme = mvarstrCriterioInforme & " AND {tipos_reactivo_ex.nombre} like " & Chr(34) & "*" & Trim(txtFiltro(1)) & "*" & Chr(34)
    End If
    ' Lote
    If Trim(Trim(txtFiltro(2))) <> "" Then
        filtro = filtro & " AND be.LOTE like '%" & Trim(txtFiltro(2)) & "%'"
        mvarstrCriterioInforme = mvarstrCriterioInforme & " AND {botes_ex.LOTE} like " & Chr(34) & "*" & Trim(txtFiltro(2)) & "*" & Chr(34)
    End If
    If chkProbetas.Value = Checked Then
        filtro = filtro & " AND tr.PROBETA = 1 "
        mvarstrCriterioInforme = mvarstrCriterioInforme & " AND {tipos_reactivo_ex.PROBETA} = 1 "
    End If
    If cmbCentro.Text <> "" Then
        filtro = filtro & " AND be.CENTRO_ID = " & cmbCentro.BoundText
        mvarstrCriterioInforme = mvarstrCriterioInforme & " AND {botes_ex.CENTRO_ID} = " & cmbCentro.BoundText
    End If
    ' Query
    ' M1136 : Campo NO_CADUCA
    ' M2642 : Filtro por proveedores
    If cmbProveedores.getTEXTO <> "" Then
        filtro = filtro & " AND pbe.PROVEEDOR_ID = " & cmbProveedores.getPK_SALIDA
    End If
    If opHenkel(0).Value = False Then
        If opHenkel(1).Value = True Then
            filtro = filtro & " AND be.HENKEL = 1 "
        Else
            filtro = filtro & " AND be.HENKEL = 0 "
        End If
    End If
        
    consulta = "SELECT be.id_bote_ex, " & _
               "       tb.codigo, " & _
               "       tr.nombre, " & _
               "       be.fecha_recepcion, " & _
               "       be.fecha_apertura, " & _
               "       be.fecha_fin, " & _
               "       be.fecha_caducidad, " & _
               "       be.tipo_bote_ex_id, " & _
               "       be.LOTE, " & _
               "       tb.precio, " & _
               "       tb.cantidad, tb.tipo_m_referencia_id, be.numero,be.codigo,be.no_conforme,be.NO_CADUCA,c.NOMBRE,be.certificado " & _
               " FROM BOTES_EX be, TIPOS_BOTE_EX tb, TIPOS_REACTIVO_EX tr, centros C " & _
               IIf(cmbProveedores.getTEXTO <> "", ",pedidos_bote_ex pbe", "") & _
               " WHERE be.tipo_bote_ex_id = tb.id_tipo_bote_ex AND be.CENTRO_ID = c.ID_CENTRO " & _
               "   AND tb.tipo_reactivo_ex_id = tr.id_tipo_reactivo_ex " & _
                   strReactivo & strBote & fecha_desde & fecha_hasta & strAbierto & _
                   strTerminado & strCaducado & strconforme & strAnulado & matref & _
                IIf(cmbProveedores.getTEXTO <> "", " and be.PEDIDO_BOTE_EX_ID = pbe.ID_PEDIDO_BOTE_EX and be.tipo_bote_ex_id = pbe.tipo_bote_ex_id ", "") & _
                   filtro & _
               " ORDER BY be.id_bote_ex desc"
    Me.MousePointer = 11
    Set rs = datos_bd(consulta)
    If Trim(mvarstrCriterioInforme) <> "" Then mvarstrCriterioInforme = Mid(mvarstrCriterioInforme, 5)
    If rs.RecordCount > 0 Then
        While Not rs.EOF
            With lista.ListItems.Add(, , Format(rs.Fields(0), "00000"))
                .SubItems(1) = rs(13) & "-" & Format(rs(12), "000") & "-" & Format(rs(3), "yy") ' Número particular
                .SubItems(2) = rs(10)
                .SubItems(3) = rs.Fields(2)
                If Not IsNull(rs.Fields(3)) Then
                    .SubItems(4) = rs.Fields(3)
                End If
                If Not IsNull(rs.Fields(4)) Then
                    .SubItems(5) = rs.Fields(4)
                End If
                If Not IsNull(rs.Fields(5)) Then
                    .SubItems(6) = rs.Fields(5)
                End If
                If IsNull(rs.Fields(6)) Then
                    'M1136-I
                    .SubItems(7) = "N.A."
                    'M1136-F
                Else
                    'M1136-I
                    If rs(15) = 1 Then
                        .SubItems(7) = "N.A."
                    Else
                        .SubItems(7) = rs.Fields(6)
                    End If
                    'M1136-F
                End If
                .SubItems(8) = rs(7)
                .SubItems(9) = rs(8)
                .SubItems(10) = Format(rs(9), "currency")
                .SubItems(11) = rs(11) ' TIPO DE REACTIVO
                .SubItems(12) = rs(16) ' CENTRO_ID
                ' Pelota de colores según la evaluación del reactivo
                If .SubItems(11) = "2" Or _
                   .SubItems(11) = "3" Or _
                   .SubItems(11) = "6" Then
                   .SmallIcon = CInt(rs(17)) + 1
                End If
                ' Si no conforme coloreamos la linea
                If rs(14) = 1 Then
                    lista_colorear lista, lista.ListItems.Count, vbRed
                End If
            
            End With
            rs.MoveNext
        Wend
        lblMsg.Caption = "Botes entre el " & Format(fdesde, "dd/mm/yyyy") & " y " & Format(fhasta, "dd/mm/yyyy") & " (Encontrados : " & rs.RecordCount & ")"
        lista_Click
    Else
        lblMsg.Caption = "No existe ningun bote con esos criterios."
    End If
    Me.MousePointer = 0
    Set rs = Nothing
    Exit Sub
fallo:
    Me.MousePointer = 0
    MsgBox "Se ha producido un error al buscar los Botes.", vbCritical, Err.Description
End Sub

Private Sub lista_Click()
    cmdEtiqueta.Enabled = False
    cmdAbrir.Enabled = False
    cmdTerminar.Enabled = False
    cmdModificar.Enabled = False
    cmdPanreac.Enabled = False
' Cambio solicitado por Jennifer, poder adjuntar certificado externo a cualquier reactivo
'    cmdCertificadoExterno.Enabled = False
    cmdCertificadoExterno.Enabled = True
    Command2.Enabled = False
    If lista.ListItems.Count > 0 Then
        cmdModificar.Enabled = True
'        If opTipo(0).value = True Then
            cmdEtiqueta.Enabled = True
            If lista.ListItems(lista.selectedItem.Index).SubItems(5) = "" Then
                cmdTerminar.Enabled = False
                cmdAbrir.Enabled = True
            Else
                If lista.ListItems(lista.selectedItem.Index).SubItems(6) = "" Then
                    cmdTerminar.Enabled = True
                    cmdAbrir.Enabled = False
                End If
            End If
'        End If
        ' Material certificado, activa los botones
        If lista.ListItems(lista.selectedItem.Index).SubItems(11) = "2" Or _
           lista.ListItems(lista.selectedItem.Index).SubItems(11) = "3" Or _
           lista.ListItems(lista.selectedItem.Index).SubItems(11) = "6" Then
            cmdPanreac.Enabled = True
            cmdCertificadoExterno.Enabled = True
            Command2.Enabled = True
        End If
    End If
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

Private Sub lista_DblClick()
    If lista.ListItems.Count > 0 Then
'        gbotereactivoex = lista.ListItems(lista.SelectedItem.Index).SubItems(7)
        frmREX_Bote.PK = lista.ListItems(lista.selectedItem.Index).SubItems(8)
        frmREX_Bote.Show 1
'        gbotereactivoex = 0
    End If
End Sub

Private Sub actualizar_lista()
'    If gbotereactivoex > 0 Then
    If lista.ListItems.Count > 0 Then
        Dim oBote As New clsBotes_ex
        With oBote
'            If .CARGAR(gbotereactivoex) = True Then
            If .CARGAR(CLng(lista.ListItems(lista.selectedItem.Index).Text)) Then
                Dim oTb As New clsTipos_bote_ex
                oTb.CARGAR oBote.getTIPO_BOTE_EX_ID
                If oTb.getTIPO_M_REFERENCIA_ID = 2 Or oTb.getTIPO_M_REFERENCIA_ID = 3 Or oTb.getTIPO_M_REFERENCIA_ID = 6 Then
                    lista.ListItems(lista.selectedItem.Index).SmallIcon = CInt(oBote.getCERTIFICADO) + 1
                End If
                
                lista.ListItems(lista.selectedItem.Index).SubItems(5) = Format(.getFECHA_APERTURA, "dd/mm/yyyy")
                lista.ListItems(lista.selectedItem.Index).SubItems(6) = Format(.getFECHA_FIN, "dd/mm/yyyy")
                'M1136-I
                If .getNO_CADUCA = 1 Then
                    lista.ListItems(lista.selectedItem.Index).SubItems(7) = "N.A."
                Else
                    lista.ListItems(lista.selectedItem.Index).SubItems(7) = Format(.getFECHA_CADUCIDAD, "dd/mm/yyyy")
                End If
                'M1136-F
                lista.ListItems(lista.selectedItem.Index).SubItems(9) = .getLOTE
                
                ' Si no conforme coloreamos la linea
                If .getNO_CONFORME = 1 Then
                    lista_colorear lista, lista.selectedItem.Index, vbRed
                Else
                    lista_colorear lista, lista.selectedItem.Index, vbBlack
                End If
                
            End If
        End With
        Set oBote = Nothing
        lista_Click
    End If
End Sub
'Private Sub cmdImprimir_Click_old()
'    On Error GoTo fallo
'    Dim i As Integer
'    Dim total As Currency
'    ' Generamos los datos del listado
'    Dim rs As New ADODB.RecordSet
'    rs.Fields.Append "c1", adChar, 15, adFldUpdatable
'    rs.Fields.Append "c2", adChar, 150, adFldUpdatable
'    rs.Fields.Append "c3", adChar, 150, adFldUpdatable
'    rs.Fields.Append "c4", adChar, 150, adFldUpdatable
'    rs.Fields.Append "c5", adChar, 15, adFldUpdatable ' Precio
'    rs.Open
'    For i = 1 To lista.ListItems.Count
'        rs.AddNew
'        If Trim(lista.ListItems(i).SubItems(1)) <> "" Then
'            rs("c1") = lista.ListItems(i).SubItems(1)
'        Else
'            rs("c1") = lista.ListItems(i).Text
'        End If
'        If Trim(lista.ListItems(i).SubItems(3)) <> "" Then
'            rs("c2") = lista.ListItems(i).SubItems(3)
'        End If
'        If Trim(lista.ListItems(i).SubItems(9)) <> "" Then
'            rs("c3") = lista.ListItems(i).SubItems(9)
'        End If
'        If Trim(lista.ListItems(i).SubItems(2)) <> "" Then
'            rs("c4") = lista.ListItems(i).SubItems(2)
'        End If
'        If Trim(lista.ListItems(i).SubItems(10)) <> "" Then
'            rs("c5") = lista.ListItems(i).SubItems(10)
'            total = total + lista.ListItems(i).SubItems(10)
'        End If
'        rs.Update
'    Next
'    ' Generar Listado
'    Dim Listado As New rptListadoReactivos
'    ' Cabecera
'    With Listado.Sections("cabecera")
'        .Controls("titulo").Caption = "Listado de Botes de Reactivos"
'        .Controls("etiqueta4").Caption = "Número"
'        .Controls("etiqueta5").Caption = "Reactivo"
'        .Controls("etiqueta10").Caption = "Lote"
'        .Controls("etiqueta11").Caption = "Cantidad"
'    End With
'    Set Listado.Sections("cabecera").Controls("logoc").Picture = LoadPicture(ReadINI(App.Path + "\config.ini", "logo", "logo"))
'    'Detalle
'    With Listado.Sections("detalle")
'        .Controls("d1").DataField = rs.Fields("c1").Name
'        .Controls("d2").DataField = rs.Fields("c2").Name
'        .Controls("d3").DataField = rs.Fields("c3").Name
'        .Controls("d4").DataField = rs.Fields("c4").Name
'        .Controls("d5").DataField = rs.Fields("c5").Name
'    End With
'    ' Pie de Pagina
'    With Listado.Sections("pie")
'        .Controls("pie1").Caption = "Fecha : " & Format(Date, "dd-mm-yyyy")
'        .Controls("pie2").Caption = "Impreso por : " & USUARIO.getNOMBRE
'    End With
'    With Listado.Sections("totales")
'        .Controls("LBLT1").Caption = Format(total, "CURRENCY")
'    End With
'    Set Listado.DataSource = rs
'    Listado.Caption = "Listado de Botes de Reactivos"
'    Listado.Show
'    Set rs = Nothing
'    Exit Sub
'fallo:
'    MsgBox "Error al generar el listado.", vbCritical, Err.Description
'End Sub
'
Private Sub lista_KeyUp(KeyCode As Integer, Shift As Integer)
    lista_Click
End Sub

Private Sub opAbierto_Click(Index As Integer)
    cmdBuscar_Click
End Sub

Private Sub opCaducado_Click(Index As Integer)
    cmdBuscar_Click
End Sub

Private Sub opHenkel_Click(Index As Integer)
    cmdBuscar_Click
End Sub

Private Sub opNoConforme_Click(Index As Integer)
    cmdBuscar_Click
End Sub

Private Sub opTerminado_Click(Index As Integer)
    cmdBuscar_Click
End Sub

Private Sub opTipo_Click(Index As Integer)
    cmdBuscar_Click
End Sub

Private Sub txtcodigo_GotFocus()
    txtCodigo.BackColor = &H80C0FF
    txtCodigo.SelStart = 0
    txtCodigo.SelLength = Len(txtCodigo)
End Sub
Private Sub txtcodigo_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 And txtCodigo <> "" Then
        SendKeys "{Tab}", True
        KeyAscii = 0 ' Para evitar el "bip" del sistema
    End If
End Sub

Private Sub txtcodigo_LostFocus()
    txtCodigo.BackColor = &HFFFFFF
    CARGAR_CODIGO (txtCodigo)
'    txtcodigo = ""
End Sub


Public Sub CARGAR_CODIGO(CODIGO As String)
    On Error GoTo fallo
    Dim consulta As String
    If CODIGO <> "" Then
        lista.ListItems.Clear
        Dim rs As ADODB.Recordset
        ' Query
        consulta = "SELECT be.id_bote_ex, " & _
                   "       tb.codigo, " & _
                   "       tr.nombre, " & _
                   "       be.fecha_recepcion, " & _
                   "       be.fecha_apertura, " & _
                   "       be.fecha_fin, " & _
                   "       be.fecha_caducidad, " & _
                   "       be.tipo_bote_ex_id, " & _
                   "       be.LOTE, " & _
                   "       tb.precio, " & _
                   "       tb.cantidad,tb.tipo_m_referencia_id, be.numero,be.codigo,be.no_caduca,c.nombre,be.certificado " & _
                   " FROM BOTES_EX be, " & _
                   "      TIPOS_BOTE_EX tb, " & _
                   "      TIPOS_REACTIVO_EX tr, CENTROS c " & _
                   " WHERE be.tipo_bote_ex_id = tb.id_tipo_bote_ex AND be.CENTRO_ID = c.ID_CENTRO " & _
                   "   AND tb.tipo_reactivo_ex_id = tr.id_tipo_reactivo_ex " & _
                   "   AND be.id_bote_ex = " & CLng(CODIGO) & _
                   " ORDER BY be.id_bote_ex desc"
        Me.MousePointer = 11
        Set rs = datos_bd(consulta)
        If rs.RecordCount > 0 Then
            While Not rs.EOF
                With lista.ListItems.Add(, , Format(rs.Fields(0), "00000"))
                    If rs(12) <> 0 Then
                        .SubItems(1) = rs(13) & "-" & Format(rs(12), "000") & "-" & Format(rs(3), "yy") ' Número particular
                    End If
                    .SubItems(2) = rs(10)
                    If IsNull(rs.Fields(2)) Then
                        .SubItems(3) = ""
                    Else
                        .SubItems(3) = rs.Fields(2)
                    End If
                    If IsNull(rs.Fields(3)) Then
                        .SubItems(4) = ""
                    Else
                        .SubItems(4) = rs.Fields(3)
                    End If
                    If IsNull(rs.Fields(4)) Then
                        .SubItems(5) = ""
                    Else
                        .SubItems(5) = rs.Fields(4)
                    End If
                    If IsNull(rs.Fields(5)) Then
                        .SubItems(6) = ""
                    Else
                        .SubItems(6) = rs.Fields(5)
                    End If
                    If IsNull(rs.Fields(6)) Then
                        'M1136-I
                        .SubItems(7) = "N.A."
                        'M1136-I
                    Else
                        'M1136-I
                        If rs(14) = 1 Then
                            .SubItems(7) = "N.A."
                        Else
                            .SubItems(7) = rs.Fields(6)
                        End If
                        'M1136-F
                    End If
                    .SubItems(8) = rs(7)
                    .SubItems(9) = rs(8)
                    .SubItems(10) = Format(rs(9), "currency")
                    .SubItems(11) = rs(11)
                    .SubItems(12) = rs(15) ' CENTRO_ID
                    ' Pelota de colores según la evaluación del reactivo
                    If .SubItems(11) = "2" Or _
                       .SubItems(11) = "3" Or _
                       .SubItems(11) = "6" Then
                       .SmallIcon = CInt(rs(16)) + 1
                    End If
                    ' Si no conforme coloreamos la linea
                    If rs(14) = 1 Then
                        lista_colorear lista, lista.ListItems.Count, vbRed
                    End If
                        
                    End With
                rs.MoveNext
            Wend
            lblMsg.Caption = "Bote localizado (Encontrados : " & rs.RecordCount & ")"
            lista_Click
        Else
            lblMsg.Caption = "No existe ningun bote con esos criterios."
        End If
    End If
    Me.MousePointer = 0
    Set rs = Nothing
    Exit Sub
fallo:
    Me.MousePointer = 0
    MsgBox "Se ha producido un error al buscar los Botes.", vbCritical, Err.Description
End Sub
Private Sub nueva_etiqueta()
    Dim cadena As String
   On Error GoTo nueva_etiqueta_Error

    For i = 1 To lista.ListItems.Count
        If lista.ListItems(i).Checked = True Then
            cadena = cadena & lista.ListItems(i).Text & ","
        End If
    Next
    If cadena <> "" Then
        Dim oBote As New clsBotes_ex
        oBote.imprimir_etiqueta Left(cadena, Len(cadena) - 1)
    Else
        MsgBox "Marque los botes para los que desea generar etiquetas.", vbExclamation, App.Title
    End If

   On Error GoTo 0
   Exit Sub

nueva_etiqueta_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure nueva_etiqueta of Formulario frmREX_Gestion"
End Sub

Private Sub txtfiltro_Change(Index As Integer)
    cmdBuscar_Click
End Sub

Private Sub txtfiltro_GotFocus(Index As Integer)
    txtFiltro(Index).SelStart = 0
    txtFiltro(Index).SelLength = Len(txtFiltro(Index))
    txtFiltro(Index).BackColor = &HC0FFFF
End Sub

Private Sub txtfiltro_LostFocus(Index As Integer)
    txtFiltro(Index).BackColor = vbWhite

End Sub

Private Sub cargar_reactivos()
    ' Mat. Referencia
    cmbReactivos.limpiar
    Dim matref As String
    Dim h As Integer
    Dim aux As String
    For h = 0 To 6
        If chktiporeactivo(h).Value = Checked Then
            aux = aux & h + 1 & ","
        End If
    Next
    If Len(aux) > 0 Then
        matref = " AND TB.tipo_m_referencia_id in (" & Left(aux, Len(aux) - 1) & ")"
    End If
    Dim consulta As String
    consulta = " SELECT DISTINCT T.ID_TIPO_REACTIVO_EX,T.NOMBRE " & _
               "   FROM TIPOS_REACTIVO_EX T, TIPOS_BOTE_EX TB " & _
               "  WHERE T.ID_TIPO_REACTIVO_EX = TB.TIPO_REACTIVO_EX_ID " & _
               matref
    Dim conn As ADODB.Connection
    If CrearConexionGlobal(conn, "", "") = True Then ' CONECTAR EL RS
        With cmbReactivos
            .setCONN = conn
            .setFK_CAMPO = ""
            .setFK_VALOR = 0
            .setTABLA = "TIPOS_REACTIVO_EX"
            .setDESCRIPCION = "Sustancias / Materiales"
            .setPK = "T.ID_TIPO_REACTIVO_EX"
            .setCAMPO = "T.NOMBRE"
            .setFILTRO = ""
            .setQUERY = consulta
            .setMUESTRA_DETALLE = True
            Set .FORMULARIO = frmREX_Reactivo
        End With
    End If
End Sub
