VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmEquipoPuestaServicio 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Datos de Recepción"
   ClientHeight    =   9930
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   9600
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9930
   ScaleWidth      =   9600
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame frmUsuario 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   "Usuario"
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
      Height          =   960
      Left            =   45
      TabIndex        =   64
      Top             =   8910
      Visible         =   0   'False
      Width           =   3660
      Begin VB.TextBox texto 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         Height          =   285
         Index           =   0
         Left            =   990
         Locked          =   -1  'True
         MaxLength       =   100
         TabIndex        =   65
         Top             =   225
         Width           =   2490
      End
      Begin MSComCtl2.DTPicker txtFechaRecepcion 
         Height          =   345
         Left            =   990
         TabIndex        =   68
         Top             =   540
         Width           =   1410
         _ExtentX        =   2487
         _ExtentY        =   609
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
         Format          =   52494337
         CurrentDate     =   2
         MinDate         =   2
      End
      Begin VB.Label Label2 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Caption         =   "Fecha"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   6
         Left            =   135
         TabIndex        =   67
         Top             =   585
         Width           =   825
      End
      Begin VB.Label Label2 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Caption         =   "Usuario :"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   0
         Left            =   135
         TabIndex        =   66
         Top             =   270
         Width           =   825
      End
   End
   Begin VB.Frame Frame3 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   "Documentación que Acompaña"
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
      Height          =   1725
      Left            =   45
      TabIndex        =   18
      Top             =   3870
      Width           =   9510
      Begin VB.Frame Frame9 
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   0  'None
         Height          =   465
         Left            =   5310
         TabIndex        =   52
         Top             =   855
         Visible         =   0   'False
         Width           =   3975
         Begin VB.OptionButton op6 
            BackColor       =   &H00C0C0C0&
            Caption         =   "No Aplica"
            Height          =   285
            Index           =   2
            Left            =   2700
            TabIndex        =   55
            Top             =   135
            Width           =   1230
         End
         Begin VB.OptionButton op6 
            BackColor       =   &H00C0C0C0&
            Caption         =   "No"
            Height          =   285
            Index           =   1
            Left            =   1395
            TabIndex        =   54
            Top             =   135
            Width           =   1230
         End
         Begin VB.OptionButton op6 
            BackColor       =   &H00C0C0C0&
            Caption         =   "Si"
            Height          =   285
            Index           =   0
            Left            =   180
            TabIndex        =   53
            Top             =   135
            Width           =   1050
         End
      End
      Begin VB.Frame Frame8 
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   0  'None
         Height          =   465
         Left            =   5310
         TabIndex        =   48
         Top             =   495
         Visible         =   0   'False
         Width           =   3975
         Begin VB.OptionButton op5 
            BackColor       =   &H00C0C0C0&
            Caption         =   "Si"
            Height          =   285
            Index           =   0
            Left            =   180
            TabIndex        =   51
            Top             =   135
            Width           =   1050
         End
         Begin VB.OptionButton op5 
            BackColor       =   &H00C0C0C0&
            Caption         =   "No"
            Height          =   285
            Index           =   1
            Left            =   1395
            TabIndex        =   50
            Top             =   135
            Width           =   1230
         End
         Begin VB.OptionButton op5 
            BackColor       =   &H00C0C0C0&
            Caption         =   "No Aplica"
            Height          =   285
            Index           =   2
            Left            =   2700
            TabIndex        =   49
            Top             =   135
            Width           =   1230
         End
      End
      Begin VB.TextBox txtop7 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   1530
         MaxLength       =   100
         TabIndex        =   25
         Top             =   1350
         Width           =   7755
      End
      Begin VB.Frame Frame7 
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   0  'None
         Height          =   465
         Left            =   5310
         TabIndex        =   44
         Top             =   135
         Visible         =   0   'False
         Width           =   3975
         Begin VB.OptionButton op4 
            BackColor       =   &H00C0C0C0&
            Caption         =   "No Aplica"
            Height          =   285
            Index           =   2
            Left            =   2700
            TabIndex        =   47
            Top             =   135
            Width           =   1230
         End
         Begin VB.OptionButton op4 
            BackColor       =   &H00C0C0C0&
            Caption         =   "No"
            Height          =   285
            Index           =   1
            Left            =   1395
            TabIndex        =   46
            Top             =   135
            Width           =   1230
         End
         Begin VB.OptionButton op4 
            BackColor       =   &H00C0C0C0&
            Caption         =   "Si"
            Height          =   285
            Index           =   0
            Left            =   180
            TabIndex        =   45
            Top             =   135
            Width           =   1050
         End
      End
      Begin VB.Label lblop 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Caption         =   "Otros"
         ForeColor       =   &H00000000&
         Height          =   285
         Index           =   7
         Left            =   180
         TabIndex        =   24
         Top             =   1395
         Width           =   1995
      End
      Begin VB.Label lblop 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Caption         =   "Certificado de Calibración"
         ForeColor       =   &H00000000&
         Height          =   285
         Index           =   6
         Left            =   180
         TabIndex        =   23
         Top             =   1035
         Width           =   5280
      End
      Begin VB.Label lblop 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Caption         =   "Certificado de Fabricación"
         ForeColor       =   &H00000000&
         Height          =   285
         Index           =   5
         Left            =   180
         TabIndex        =   22
         Top             =   675
         Width           =   5280
      End
      Begin VB.Label lblop 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Caption         =   "Manual del Fabricante"
         ForeColor       =   &H00000000&
         Height          =   285
         Index           =   4
         Left            =   180
         TabIndex        =   19
         Top             =   315
         Width           =   5280
      End
   End
   Begin VB.Frame Frame6 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   "Observaciones"
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
      Height          =   1050
      Left            =   45
      TabIndex        =   30
      Top             =   7020
      Width           =   9510
      Begin VB.TextBox txtob 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   735
         Left            =   180
         MaxLength       =   100
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   31
         Top             =   225
         Width           =   9105
      End
   End
   Begin VB.Frame Frame5 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   "Resultado de la inspección"
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
      Height          =   780
      Left            =   45
      TabIndex        =   21
      Top             =   8100
      Width           =   9510
      Begin VB.OptionButton opCONFORME 
         BackColor       =   &H00C0C0C0&
         Caption         =   "NO CONFORME"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   1
         Left            =   4635
         TabIndex        =   63
         Top             =   315
         Width           =   2850
      End
      Begin VB.OptionButton opCONFORME 
         BackColor       =   &H00C0C0C0&
         Caption         =   "CONFORME"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   0
         Left            =   1620
         TabIndex        =   62
         Top             =   315
         Width           =   2490
      End
   End
   Begin VB.Frame Frame4 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   "Inspección Previa"
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
      Height          =   1365
      Left            =   45
      TabIndex        =   20
      Top             =   5625
      Width           =   9510
      Begin VB.Frame Frame11 
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   0  'None
         Height          =   465
         Left            =   5310
         TabIndex        =   59
         Top             =   495
         Visible         =   0   'False
         Width           =   3975
         Begin VB.OptionButton op9 
            BackColor       =   &H00C0C0C0&
            Caption         =   "No"
            Height          =   285
            Index           =   1
            Left            =   1395
            TabIndex        =   61
            Top             =   135
            Width           =   1230
         End
         Begin VB.OptionButton op9 
            BackColor       =   &H00C0C0C0&
            Caption         =   "Si"
            Height          =   285
            Index           =   0
            Left            =   180
            TabIndex        =   60
            Top             =   135
            Width           =   1050
         End
      End
      Begin VB.Frame Frame10 
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   0  'None
         Height          =   465
         Left            =   5310
         TabIndex        =   56
         Top             =   135
         Visible         =   0   'False
         Width           =   3975
         Begin VB.OptionButton op8 
            BackColor       =   &H00C0C0C0&
            Caption         =   "Si"
            Height          =   285
            Index           =   0
            Left            =   180
            TabIndex        =   58
            Top             =   135
            Width           =   1050
         End
         Begin VB.OptionButton op8 
            BackColor       =   &H00C0C0C0&
            Caption         =   "No"
            Height          =   285
            Index           =   1
            Left            =   1395
            TabIndex        =   57
            Top             =   135
            Width           =   1230
         End
      End
      Begin VB.TextBox txtop10 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   1485
         MaxLength       =   100
         TabIndex        =   29
         Top             =   990
         Width           =   7800
      End
      Begin VB.Label lblop 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Caption         =   "Otros"
         ForeColor       =   &H00000000&
         Height          =   285
         Index           =   10
         Left            =   135
         TabIndex        =   28
         Top             =   990
         Width           =   1995
      End
      Begin VB.Label lblop 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Caption         =   "Vienen todos los accesorios descritos en el Manual"
         ForeColor       =   &H00000000&
         Height          =   285
         Index           =   9
         Left            =   135
         TabIndex        =   27
         Top             =   630
         Width           =   5280
      End
      Begin VB.Label lblop 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Caption         =   "Presenta Daño Exterior"
         ForeColor       =   &H00000000&
         Height          =   285
         Index           =   8
         Left            =   135
         TabIndex        =   26
         Top             =   270
         Width           =   5280
      End
   End
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Enabled         =   0   'False
      ForeColor       =   &H80000008&
      Height          =   1455
      Left            =   45
      TabIndex        =   5
      Top             =   585
      Width           =   9510
      Begin VB.TextBox texto 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         Height          =   285
         Index           =   18
         Left            =   6795
         Locked          =   -1  'True
         MaxLength       =   100
         TabIndex        =   13
         Top             =   135
         Width           =   2535
      End
      Begin VB.TextBox texto 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         Height          =   285
         Index           =   16
         Left            =   1710
         Locked          =   -1  'True
         MaxLength       =   100
         TabIndex        =   11
         Top             =   1080
         Width           =   7620
      End
      Begin VB.TextBox texto 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         Height          =   285
         Index           =   21
         Left            =   1710
         Locked          =   -1  'True
         MaxLength       =   100
         TabIndex        =   2
         Top             =   765
         Width           =   7620
      End
      Begin VB.TextBox texto 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         Height          =   285
         Index           =   20
         Left            =   1710
         Locked          =   -1  'True
         MaxLength       =   100
         TabIndex        =   1
         Top             =   450
         Width           =   7620
      End
      Begin VB.TextBox texto 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         Height          =   285
         Index           =   19
         Left            =   1710
         Locked          =   -1  'True
         MaxLength       =   100
         TabIndex        =   0
         Top             =   135
         Width           =   1770
      End
      Begin VB.Label Label2 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Nº Serie : "
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   5
         Left            =   6030
         TabIndex        =   14
         Top             =   180
         Width           =   720
      End
      Begin VB.Label Label2 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Caption         =   "Marca : "
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   1
         Left            =   135
         TabIndex        =   12
         Top             =   1080
         Width           =   1140
      End
      Begin VB.Label Label2 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Caption         =   "Inventario : "
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   4
         Left            =   135
         TabIndex        =   9
         Top             =   765
         Width           =   1320
      End
      Begin VB.Label Label2 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Caption         =   "Nombre del Equipo : "
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   3
         Left            =   135
         TabIndex        =   8
         Top             =   495
         Width           =   1635
      End
      Begin VB.Label Label2 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Caption         =   "Número del Equipo :"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   2
         Left            =   135
         TabIndex        =   7
         Top             =   180
         Width           =   1545
      End
   End
   Begin VB.CommandButton cmdok 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Aceptar"
      Height          =   870
      Left            =   7425
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   9000
      Width           =   1050
   End
   Begin VB.CommandButton cmdcancel 
      BackColor       =   &H00E0E0E0&
      Caption         =   "ESC-Salir"
      Height          =   870
      Left            =   8505
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   9000
      Width           =   1050
   End
   Begin VB.Frame Frame2 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   "Procedencia del Equipo"
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
      Height          =   1770
      Left            =   45
      TabIndex        =   6
      Top             =   2070
      Width           =   9510
      Begin VB.Frame frm3 
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   0  'None
         Height          =   600
         Left            =   5310
         TabIndex        =   40
         Top             =   1080
         Visible         =   0   'False
         Width           =   3975
         Begin VB.OptionButton op3 
            BackColor       =   &H00C0C0C0&
            Caption         =   "Conforme"
            Height          =   285
            Index           =   0
            Left            =   180
            TabIndex        =   43
            Top             =   225
            Width           =   1050
         End
         Begin VB.OptionButton op3 
            BackColor       =   &H00C0C0C0&
            Caption         =   "No Conforme"
            Height          =   285
            Index           =   1
            Left            =   1395
            TabIndex        =   42
            Top             =   225
            Width           =   1230
         End
         Begin VB.OptionButton op3 
            BackColor       =   &H00C0C0C0&
            Caption         =   "No Aplica"
            Height          =   285
            Index           =   2
            Left            =   2700
            TabIndex        =   41
            Top             =   225
            Width           =   1230
         End
      End
      Begin VB.Frame frm2 
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   0  'None
         Height          =   600
         Left            =   5310
         TabIndex        =   36
         Top             =   540
         Visible         =   0   'False
         Width           =   3975
         Begin VB.OptionButton op2 
            BackColor       =   &H00C0C0C0&
            Caption         =   "No Aplica"
            Height          =   285
            Index           =   2
            Left            =   2700
            TabIndex        =   39
            Top             =   225
            Width           =   1230
         End
         Begin VB.OptionButton op2 
            BackColor       =   &H00C0C0C0&
            Caption         =   "No Conforme"
            Height          =   285
            Index           =   1
            Left            =   1395
            TabIndex        =   38
            Top             =   225
            Width           =   1230
         End
         Begin VB.OptionButton op2 
            BackColor       =   &H00C0C0C0&
            Caption         =   "Conforme"
            Height          =   285
            Index           =   0
            Left            =   180
            TabIndex        =   37
            Top             =   225
            Width           =   1050
         End
      End
      Begin VB.OptionButton op1 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Otro"
         Height          =   285
         Index           =   3
         Left            =   8010
         TabIndex        =   35
         Top             =   270
         Width           =   1275
      End
      Begin VB.OptionButton op1 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Prestamo"
         Height          =   285
         Index           =   2
         Left            =   6705
         TabIndex        =   34
         Top             =   270
         Width           =   1275
      End
      Begin VB.OptionButton op1 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Reparación"
         Height          =   285
         Index           =   1
         Left            =   5490
         TabIndex        =   33
         Top             =   270
         Width           =   1275
      End
      Begin VB.OptionButton op1 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Compra"
         Height          =   285
         Index           =   0
         Left            =   4320
         TabIndex        =   32
         Top             =   270
         Width           =   960
      End
      Begin VB.Label lblop 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Caption         =   "PRESTAMO. Conformidad con los requisitos del ensayo en cuanto a incertidumbre, trazabilidad y mantenimiento."
         ForeColor       =   &H00000000&
         Height          =   420
         Index           =   3
         Left            =   135
         TabIndex        =   17
         Top             =   1260
         Visible         =   0   'False
         Width           =   5280
      End
      Begin VB.Label lblop 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Caption         =   "Procedencia del Equipo"
         ForeColor       =   &H00000000&
         Height          =   285
         Index           =   1
         Left            =   135
         TabIndex        =   16
         Top             =   315
         Width           =   2040
      End
      Begin VB.Label lblop 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Caption         =   "COMPRA. Conformidad con las especificaciones de la ""Hoja de pedido a Proveedores""."
         ForeColor       =   &H00000000&
         Height          =   465
         Index           =   2
         Left            =   135
         TabIndex        =   15
         Top             =   765
         Visible         =   0   'False
         Width           =   5100
      End
   End
   Begin VB.Image imagen 
      Height          =   480
      Left            =   9045
      Picture         =   "frmEquipoPuestaServicio.frx":0000
      Top             =   15
      Width           =   480
   End
   Begin VB.Label lbltitulo 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "Datos de Recepción"
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
      TabIndex        =   10
      Top             =   120
      Width           =   8640
      WordWrap        =   -1  'True
   End
   Begin VB.Shape fondo 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   555
      Left            =   0
      Top             =   0
      Width           =   9735
   End
End
Attribute VB_Name = "frmEquipoPuestaServicio"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public PK As Long
Private Sub cmdcancel_Click()
    Unload Me
End Sub

Private Sub cmdok_Click()
   On Error GoTo cmdok_Click_Error

    If validar = True Then
        Dim oEvaluacion As New clsRex_botes_certificados
        With oEvaluacion
            .setBOTE_EX_ID = BOTE_EX_ID
            .setC01_NUMERO_CERTIFICADO = texto(19)
            .setC02_ENTIDAD = texto(20)
            .setC03_INVENTARIO = texto(21)
            .setC04_ORGANISMO = texto(0)
            .setC05_TITULO_DOCUMENTO = texto(1)
            .setC06_NOMBRE_MATERIAL = texto(2)
            .setC07_FABRICANTE = texto(3)
            .setC08_FECHA_CERTIFICACION = texto(4)
            .setC09_CODIGO_LOTE = texto(5)
            .setC10_UTILIZACION_PREVISTA = texto(6)
            .setC11_INSTRUCCIONES_USO = texto(7)
            .setC12_SITUACION_PELIGROSA = texto(8)
            .setC13_NIVEL_HOMOGENEIDAD = texto(9)
            .setC14_CONCENTRACION = texto(10)
            .setC15_MATRIZ = texto(11)
            .setC16_PRESENTACION = texto(12)
            .setC17_CANTIDAD = texto(13)
            .setC18_ESTABILIDAD = texto(14)
            .setC19_TRAZABILIDAD = texto(15)
            If op(16).value = Checked Then
                .setC20_RESULTADO = 1
            Else
                .setC20_RESULTADO = 0
            End If
            .setC21_USO_PREVISTO = texto(17)
            If op(17).value = Checked Then
                .setC22_CONFORME_PROPIEDAD = 1
            Else
                .setC22_CONFORME_PROPIEDAD = 0
            End If
            .setC23_TECNICO_RESPONSABLE = usuario.getID_EMPLEADO
            .setC24_FECHA_EVALUACION = Format(Date, "dd-mm-yyyy")
            If .Insertar = 0 Then
                If .modificar(BOTE_EX_ID) = True Then
                    MsgBox "Los datos de la certificación se han modificado correctamente.", vbInformation, App.Title
                    Unload Me
                End If
            Else
                MsgBox "Los datos de la certificación se han insertado correctamente.", vbInformation, App.Title
                Unload Me
            End If
        End With
    End If

   On Error GoTo 0
   Exit Sub

cmdok_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure cmdok_Click of Formulario frmEquipoPuestaServicio"
End Sub

Private Sub Form_Load()
    log (Me.Name)
    cargar_botones Me
    If PK <> 0 Then
        cargar
    End If
'    If BOTE_EX_ID = 0 Then
'        texto(21) = BOTE_EX_ID
'    Else
'        cargar_certificado
'    End If
End Sub

Private Sub op1_Click(Index As Integer)
    lblop(2).Visible = False
    lblop(3).Visible = False
    frm2.Visible = False
    frm3.Visible = False
    If op1(0).value = True Then
        lblop(2).Visible = True
        frm2.Visible = True
    End If
    If op1(2).value = True Then
        lblop(3).Visible = True
        frm3.Visible = True
    End If
End Sub

Private Function validar() As Boolean
    validar = True
    ' Procedencia
    If op1(0).value = False And op1(1).value = False And op1(2).value = False And op1(3).value = False Then
        validar = False
    End If
    ' Compra
    If op1(0).value = True Then
        If op2(0).value = False And op2(1).value = False And op2(2).value = False Then
            validar = False
        End If
    End If
    ' Prestamo
    If op1(2).value = True Then
        If op3(0).value = False And op3(1).value = False And op3(2).value = False Then
            validar = False
        End If
    End If
    ' Documentacion
    If op4(0).value = False And op4(1).value = False And op4(2).value = False Then
        validar = False
    End If
    If op5(0).value = False And op5(1).value = False And op5(2).value = False Then
        validar = False
    End If
    If op6(0).value = False And op6(1).value = False And op6(2).value = False Then
        validar = False
    End If
    ' Inspeccion
    If op8(0).value = False And op8(1).value = False Then
        validar = False
    End If
    If op9(0).value = False And op9(1).value = False Then
        validar = False
    End If
    If validar = False Then
        MsgBox "Por favor, rellene todos las opciones.", vbExclamation, App.Title
    End If
End Function
Private Sub cargar()
    Dim oEquipo As New clsEquipos
    With oEquipo
        .Carga PK
        texto(19) = .getID_EQUIPO
        texto(18) = .getSERIE
        texto(20) = .getNOMBRE
        Dim oDeco As New clsDecodificadora
        oDeco.Carga_valor decodificadora.EQ_TIPOS_EQUIPO, .getTIPO_EQUIPO_ID
        Set oDeco = Nothing
        texto(21) = .getNOMBRE
        texto(16) = .getFABRICANTE
    End With
    Set oEquipo = Nothing
    
End Sub

Private Sub opCONFORME_Click(Index As Integer)
    If Index = 1 Then
        opCONFORME(Index).ForeColor = &HC0&
    Else
        opCONFORME(0).ForeColor = &H8000&
    End If
End Sub
