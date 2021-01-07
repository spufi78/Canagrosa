VERSION 5.00
Begin VB.Form frmFormacion_Evaluacion 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Evaluaciones"
   ClientHeight    =   9645
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10680
   Icon            =   "frmFormacion_Evaluacion.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9645
   ScaleWidth      =   10680
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdSalir 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Salir"
      Height          =   870
      Left            =   9450
      Style           =   1  'Graphical
      TabIndex        =   117
      Top             =   8640
      Width           =   1140
   End
   Begin VB.TextBox txtMedia 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0C0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   420
      Left            =   8010
      TabIndex        =   102
      Top             =   8145
      Visible         =   0   'False
      Width           =   2580
   End
   Begin VB.TextBox txtObservaciones 
      Appearance      =   0  'Flat
      Height          =   1410
      Left            =   45
      MultiLine       =   -1  'True
      TabIndex        =   98
      Top             =   8100
      Width           =   7665
   End
   Begin VB.CommandButton cmdok 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Guardar"
      Height          =   870
      Left            =   9450
      Style           =   1  'Graphical
      TabIndex        =   97
      Top             =   8640
      Width           =   1140
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Valoración"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5640
      Left            =   45
      TabIndex        =   7
      Top             =   2160
      Width           =   10590
      Begin VB.Frame Frame9 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   420
         Left            =   2385
         TabIndex        =   61
         Top             =   2970
         Width           =   7980
         Begin VB.OptionButton optAula 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            ForeColor       =   &H80000008&
            Height          =   150
            Index           =   5
            Left            =   7245
            TabIndex        =   110
            Top             =   180
            Width           =   195
         End
         Begin VB.OptionButton optAula 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            ForeColor       =   &H80000008&
            Height          =   150
            Index           =   4
            Left            =   5985
            TabIndex        =   72
            Top             =   180
            Width           =   195
         End
         Begin VB.OptionButton optAula 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            ForeColor       =   &H80000008&
            Height          =   150
            Index           =   3
            Left            =   4635
            TabIndex        =   71
            Top             =   180
            Width           =   195
         End
         Begin VB.OptionButton optAula 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            ForeColor       =   &H80000008&
            Height          =   150
            Index           =   2
            Left            =   3240
            TabIndex        =   70
            Top             =   180
            Width           =   195
         End
         Begin VB.OptionButton optAula 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            ForeColor       =   &H80000008&
            Height          =   150
            Index           =   1
            Left            =   1890
            TabIndex        =   69
            Top             =   180
            Width           =   195
         End
         Begin VB.OptionButton optAula 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            ForeColor       =   &H80000008&
            Height          =   150
            Index           =   0
            Left            =   495
            TabIndex        =   67
            Top             =   180
            Width           =   195
         End
      End
      Begin VB.Frame Frame8 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   420
         Left            =   2385
         TabIndex        =   60
         Top             =   2565
         Width           =   7980
         Begin VB.OptionButton optGrupo 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            ForeColor       =   &H80000008&
            Height          =   150
            Index           =   5
            Left            =   7245
            TabIndex        =   109
            Top             =   180
            Width           =   195
         End
         Begin VB.OptionButton optGrupo 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            ForeColor       =   &H80000008&
            Height          =   150
            Index           =   4
            Left            =   5985
            TabIndex        =   66
            Top             =   180
            Width           =   195
         End
         Begin VB.OptionButton optGrupo 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            ForeColor       =   &H80000008&
            Height          =   150
            Index           =   3
            Left            =   4635
            TabIndex        =   65
            Top             =   180
            Width           =   195
         End
         Begin VB.OptionButton optGrupo 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            ForeColor       =   &H80000008&
            Height          =   150
            Index           =   2
            Left            =   3240
            TabIndex        =   64
            Top             =   180
            Width           =   195
         End
         Begin VB.OptionButton optGrupo 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            ForeColor       =   &H80000008&
            Height          =   150
            Index           =   1
            Left            =   1890
            TabIndex        =   63
            Top             =   180
            Width           =   195
         End
         Begin VB.OptionButton optGrupo 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            ForeColor       =   &H80000008&
            Height          =   150
            Index           =   0
            Left            =   495
            TabIndex        =   62
            Top             =   180
            Width           =   195
         End
      End
      Begin VB.Frame Frame5 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   420
         Left            =   2385
         TabIndex        =   44
         Top             =   1350
         Width           =   7980
         Begin VB.OptionButton optUtilidad 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            ForeColor       =   &H80000008&
            Height          =   195
            Index           =   5
            Left            =   7245
            TabIndex        =   106
            Top             =   180
            Width           =   195
         End
         Begin VB.OptionButton optUtilidad 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            ForeColor       =   &H80000008&
            Height          =   195
            Index           =   4
            Left            =   5985
            TabIndex        =   49
            Top             =   180
            Width           =   195
         End
         Begin VB.OptionButton optUtilidad 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            ForeColor       =   &H80000008&
            Height          =   195
            Index           =   3
            Left            =   4635
            TabIndex        =   48
            Top             =   180
            Width           =   195
         End
         Begin VB.OptionButton optUtilidad 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            ForeColor       =   &H80000008&
            Height          =   195
            Index           =   2
            Left            =   3240
            TabIndex        =   47
            Top             =   180
            Width           =   195
         End
         Begin VB.OptionButton optUtilidad 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            ForeColor       =   &H80000008&
            Height          =   195
            Index           =   0
            Left            =   495
            TabIndex        =   46
            Top             =   180
            Width           =   195
         End
         Begin VB.OptionButton optUtilidad 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            ForeColor       =   &H80000008&
            Height          =   195
            Index           =   1
            Left            =   1890
            TabIndex        =   45
            Top             =   180
            Width           =   195
         End
      End
      Begin VB.Frame Frame14 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   420
         Left            =   2385
         TabIndex        =   33
         Top             =   4995
         Width           =   7980
         Begin VB.OptionButton optGeneral 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            ForeColor       =   &H0000FF00&
            Height          =   150
            Index           =   5
            Left            =   7245
            TabIndex        =   115
            Top             =   180
            Width           =   195
         End
         Begin VB.OptionButton optGeneral 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            ForeColor       =   &H0000FF00&
            Height          =   150
            Index           =   4
            Left            =   5985
            TabIndex        =   96
            Top             =   180
            Width           =   195
         End
         Begin VB.OptionButton optGeneral 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            ForeColor       =   &H0000FF00&
            Height          =   150
            Index           =   3
            Left            =   4635
            TabIndex        =   95
            Top             =   180
            Width           =   195
         End
         Begin VB.OptionButton optGeneral 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            ForeColor       =   &H0000FF00&
            Height          =   150
            Index           =   2
            Left            =   3240
            TabIndex        =   94
            Top             =   180
            Width           =   195
         End
         Begin VB.OptionButton optGeneral 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            ForeColor       =   &H0000FF00&
            Height          =   150
            Index           =   1
            Left            =   1890
            TabIndex        =   93
            Top             =   180
            Width           =   195
         End
         Begin VB.OptionButton optGeneral 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            ForeColor       =   &H0000FF00&
            Height          =   150
            Index           =   0
            Left            =   495
            TabIndex        =   92
            Top             =   180
            Width           =   195
         End
      End
      Begin VB.Frame Frame13 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   420
         Left            =   2385
         TabIndex        =   32
         Top             =   4590
         Width           =   7980
         Begin VB.OptionButton optMaterial 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            ForeColor       =   &H80000008&
            Height          =   150
            Index           =   5
            Left            =   7245
            TabIndex        =   114
            Top             =   180
            Width           =   195
         End
         Begin VB.OptionButton optMaterial 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            ForeColor       =   &H80000008&
            Height          =   150
            Index           =   4
            Left            =   5985
            TabIndex        =   91
            Top             =   180
            Width           =   195
         End
         Begin VB.OptionButton optMaterial 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            ForeColor       =   &H80000008&
            Height          =   150
            Index           =   3
            Left            =   4635
            TabIndex        =   90
            Top             =   180
            Width           =   195
         End
         Begin VB.OptionButton optMaterial 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            ForeColor       =   &H80000008&
            Height          =   150
            Index           =   2
            Left            =   3240
            TabIndex        =   89
            Top             =   180
            Width           =   195
         End
         Begin VB.OptionButton optMaterial 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            ForeColor       =   &H80000008&
            Height          =   150
            Index           =   1
            Left            =   1890
            TabIndex        =   88
            Top             =   180
            Width           =   195
         End
         Begin VB.OptionButton optMaterial 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            ForeColor       =   &H80000008&
            Height          =   150
            Index           =   0
            Left            =   495
            TabIndex        =   87
            Top             =   180
            Width           =   195
         End
      End
      Begin VB.Frame Frame12 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   420
         Left            =   2385
         TabIndex        =   31
         Top             =   4185
         Width           =   7980
         Begin VB.OptionButton optHorario 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            ForeColor       =   &H80000008&
            Height          =   150
            Index           =   5
            Left            =   7245
            TabIndex        =   113
            Top             =   180
            Width           =   195
         End
         Begin VB.OptionButton optHorario 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            ForeColor       =   &H80000008&
            Height          =   150
            Index           =   4
            Left            =   5985
            TabIndex        =   86
            Top             =   180
            Width           =   195
         End
         Begin VB.OptionButton optHorario 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            ForeColor       =   &H80000008&
            Height          =   150
            Index           =   3
            Left            =   4635
            TabIndex        =   85
            Top             =   180
            Width           =   195
         End
         Begin VB.OptionButton optHorario 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            ForeColor       =   &H80000008&
            Height          =   150
            Index           =   2
            Left            =   3240
            TabIndex        =   84
            Top             =   180
            Width           =   195
         End
         Begin VB.OptionButton optHorario 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            ForeColor       =   &H80000008&
            Height          =   150
            Index           =   1
            Left            =   1890
            TabIndex        =   83
            Top             =   180
            Width           =   195
         End
         Begin VB.OptionButton optHorario 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            ForeColor       =   &H80000008&
            Height          =   150
            Index           =   0
            Left            =   495
            TabIndex        =   82
            Top             =   180
            Width           =   195
         End
      End
      Begin VB.Frame Frame11 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   420
         Left            =   2385
         TabIndex        =   30
         Top             =   3780
         Width           =   7980
         Begin VB.OptionButton optDuracion 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            ForeColor       =   &H80000008&
            Height          =   150
            Index           =   5
            Left            =   7245
            TabIndex        =   112
            Top             =   180
            Width           =   195
         End
         Begin VB.OptionButton optDuracion 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            ForeColor       =   &H80000008&
            Height          =   150
            Index           =   4
            Left            =   5985
            TabIndex        =   81
            Top             =   180
            Width           =   195
         End
         Begin VB.OptionButton optDuracion 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            ForeColor       =   &H80000008&
            Height          =   150
            Index           =   3
            Left            =   4635
            TabIndex        =   80
            Top             =   180
            Width           =   195
         End
         Begin VB.OptionButton optDuracion 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            ForeColor       =   &H80000008&
            Height          =   150
            Index           =   2
            Left            =   3240
            TabIndex        =   79
            Top             =   180
            Width           =   195
         End
         Begin VB.OptionButton optDuracion 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            ForeColor       =   &H80000008&
            Height          =   150
            Index           =   1
            Left            =   1890
            TabIndex        =   78
            Top             =   180
            Width           =   195
         End
         Begin VB.OptionButton optDuracion 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            ForeColor       =   &H80000008&
            Height          =   150
            Index           =   0
            Left            =   495
            TabIndex        =   77
            Top             =   180
            Width           =   195
         End
      End
      Begin VB.Frame Frame10 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   420
         Left            =   2385
         TabIndex        =   29
         Top             =   3375
         Width           =   7980
         Begin VB.OptionButton optAmbiente 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            ForeColor       =   &H80000008&
            Height          =   150
            Index           =   5
            Left            =   7245
            TabIndex        =   111
            Top             =   180
            Width           =   195
         End
         Begin VB.OptionButton optAmbiente 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            ForeColor       =   &H80000008&
            Height          =   150
            Index           =   4
            Left            =   5985
            TabIndex        =   76
            Top             =   180
            Width           =   195
         End
         Begin VB.OptionButton optAmbiente 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            ForeColor       =   &H80000008&
            Height          =   150
            Index           =   3
            Left            =   4635
            TabIndex        =   75
            Top             =   180
            Width           =   195
         End
         Begin VB.OptionButton optAmbiente 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            ForeColor       =   &H80000008&
            Height          =   150
            Index           =   2
            Left            =   3240
            TabIndex        =   74
            Top             =   180
            Width           =   195
         End
         Begin VB.OptionButton optAmbiente 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            ForeColor       =   &H80000008&
            Height          =   150
            Index           =   1
            Left            =   1890
            TabIndex        =   73
            Top             =   180
            Width           =   195
         End
         Begin VB.OptionButton optAmbiente 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            ForeColor       =   &H80000008&
            Height          =   150
            Index           =   0
            Left            =   495
            TabIndex        =   68
            Top             =   180
            Width           =   195
         End
      End
      Begin VB.Frame Frame7 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   420
         Left            =   2385
         TabIndex        =   28
         Top             =   2160
         Width           =   7980
         Begin VB.OptionButton optMedios 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            ForeColor       =   &H80000008&
            Height          =   150
            Index           =   5
            Left            =   7245
            TabIndex        =   108
            Top             =   180
            Width           =   195
         End
         Begin VB.OptionButton optMedios 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            ForeColor       =   &H80000008&
            Height          =   150
            Index           =   4
            Left            =   5985
            TabIndex        =   59
            Top             =   180
            Width           =   195
         End
         Begin VB.OptionButton optMedios 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            ForeColor       =   &H80000008&
            Height          =   150
            Index           =   3
            Left            =   4635
            TabIndex        =   58
            Top             =   180
            Width           =   195
         End
         Begin VB.OptionButton optMedios 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            ForeColor       =   &H80000008&
            Height          =   150
            Index           =   2
            Left            =   3240
            TabIndex        =   57
            Top             =   180
            Width           =   195
         End
         Begin VB.OptionButton optMedios 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            ForeColor       =   &H80000008&
            Height          =   150
            Index           =   1
            Left            =   1890
            TabIndex        =   56
            Top             =   180
            Width           =   195
         End
         Begin VB.OptionButton optMedios 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            ForeColor       =   &H80000008&
            Height          =   150
            Index           =   0
            Left            =   495
            TabIndex        =   55
            Top             =   180
            Width           =   195
         End
      End
      Begin VB.Frame Frame6 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   420
         Left            =   2385
         TabIndex        =   27
         Top             =   1755
         Width           =   7980
         Begin VB.OptionButton optPracticos 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            ForeColor       =   &H80000008&
            Height          =   150
            Index           =   5
            Left            =   7245
            TabIndex        =   107
            Top             =   180
            Width           =   195
         End
         Begin VB.OptionButton optPracticos 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            ForeColor       =   &H80000008&
            Height          =   150
            Index           =   4
            Left            =   5985
            TabIndex        =   54
            Top             =   180
            Width           =   195
         End
         Begin VB.OptionButton optPracticos 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            ForeColor       =   &H80000008&
            Height          =   150
            Index           =   3
            Left            =   4635
            TabIndex        =   53
            Top             =   180
            Width           =   195
         End
         Begin VB.OptionButton optPracticos 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            ForeColor       =   &H80000008&
            Height          =   150
            Index           =   2
            Left            =   3240
            TabIndex        =   52
            Top             =   180
            Width           =   195
         End
         Begin VB.OptionButton optPracticos 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            ForeColor       =   &H80000008&
            Height          =   150
            Index           =   1
            Left            =   1890
            TabIndex        =   51
            Top             =   180
            Width           =   195
         End
         Begin VB.OptionButton optPracticos 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            ForeColor       =   &H80000008&
            Height          =   150
            Index           =   0
            Left            =   495
            TabIndex        =   50
            Top             =   180
            Width           =   195
         End
      End
      Begin VB.Frame Frame4 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   465
         Left            =   2385
         TabIndex        =   26
         Top             =   900
         Width           =   7980
         Begin VB.OptionButton optContenido 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            ForeColor       =   &H80000008&
            Height          =   195
            Index           =   5
            Left            =   7245
            TabIndex        =   105
            Top             =   225
            Width           =   195
         End
         Begin VB.OptionButton optContenido 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            ForeColor       =   &H80000008&
            Height          =   195
            Index           =   0
            Left            =   495
            TabIndex        =   43
            Top             =   225
            Width           =   195
         End
         Begin VB.OptionButton optContenido 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            ForeColor       =   &H80000008&
            Height          =   195
            Index           =   4
            Left            =   5985
            TabIndex        =   42
            Top             =   225
            Width           =   195
         End
         Begin VB.OptionButton optContenido 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            ForeColor       =   &H80000008&
            Height          =   195
            Index           =   3
            Left            =   4635
            TabIndex        =   41
            Top             =   225
            Width           =   195
         End
         Begin VB.OptionButton optContenido 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            ForeColor       =   &H80000008&
            Height          =   195
            Index           =   2
            Left            =   3240
            TabIndex        =   40
            Top             =   225
            Width           =   195
         End
         Begin VB.OptionButton optContenido 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            ForeColor       =   &H80000008&
            Height          =   195
            Index           =   1
            Left            =   1890
            TabIndex        =   39
            Top             =   225
            Width           =   195
         End
      End
      Begin VB.Frame Frame3 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   420
         Left            =   2385
         TabIndex        =   25
         Top             =   540
         Width           =   7980
         Begin VB.OptionButton optOrganizacion 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            ForeColor       =   &H0000C000&
            Height          =   195
            Index           =   0
            Left            =   495
            TabIndex        =   104
            Top             =   180
            Width           =   195
         End
         Begin VB.OptionButton optOrganizacion 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            ForeColor       =   &H0000C000&
            Height          =   195
            Index           =   5
            Left            =   7245
            TabIndex        =   38
            Top             =   180
            Width           =   195
         End
         Begin VB.OptionButton optOrganizacion 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            ForeColor       =   &H0000C000&
            Height          =   195
            Index           =   4
            Left            =   5985
            TabIndex        =   37
            Top             =   180
            Width           =   195
         End
         Begin VB.OptionButton optOrganizacion 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            ForeColor       =   &H0000C000&
            Height          =   195
            Index           =   3
            Left            =   4635
            TabIndex        =   36
            Top             =   180
            Width           =   195
         End
         Begin VB.OptionButton optOrganizacion 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            ForeColor       =   &H0000C000&
            Height          =   195
            Index           =   2
            Left            =   3240
            TabIndex        =   35
            Top             =   180
            Width           =   195
         End
         Begin VB.OptionButton optOrganizacion 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            ForeColor       =   &H0000C000&
            Height          =   195
            Index           =   1
            Left            =   1890
            TabIndex        =   34
            Top             =   180
            Width           =   195
         End
      End
      Begin VB.Label Label20 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "N/A"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   240
         Left            =   2385
         TabIndex        =   103
         Top             =   225
         Width           =   1275
      End
      Begin VB.Line Line5 
         X1              =   2340
         X2              =   10395
         Y1              =   5445
         Y2              =   5445
      End
      Begin VB.Line Line4 
         X1              =   2340
         X2              =   10395
         Y1              =   495
         Y2              =   495
      End
      Begin VB.Line Line3 
         X1              =   10395
         X2              =   10395
         Y1              =   495
         Y2              =   5445
      End
      Begin VB.Line Line1 
         X1              =   2340
         X2              =   2340
         Y1              =   495
         Y2              =   5445
      End
      Begin VB.Label Label18 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Impresión General"
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
         Height          =   240
         Left            =   135
         TabIndex        =   24
         Top             =   5130
         Width           =   2220
      End
      Begin VB.Label Label17 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Material"
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
         Height          =   240
         Left            =   135
         TabIndex        =   23
         Top             =   4725
         Width           =   2220
      End
      Begin VB.Label Label16 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Horario del curso"
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
         Height          =   240
         Left            =   135
         TabIndex        =   22
         Top             =   4320
         Width           =   2220
      End
      Begin VB.Label Label15 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Duración del curso"
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
         Height          =   240
         Left            =   135
         TabIndex        =   21
         Top             =   3915
         Width           =   2220
      End
      Begin VB.Label Label14 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Ambiente alumnado / Grupo"
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
         Height          =   240
         Left            =   135
         TabIndex        =   20
         Top             =   3510
         Width           =   2220
      End
      Begin VB.Label Label13 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Comodidad del aula"
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
         Height          =   240
         Left            =   135
         TabIndex        =   19
         Top             =   3105
         Width           =   2220
      End
      Begin VB.Label Label12 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Dinámicas de grupo"
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
         Height          =   240
         Left            =   135
         TabIndex        =   18
         Top             =   2700
         Width           =   2220
      End
      Begin VB.Label Label11 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Medios audiovisuales"
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
         Height          =   240
         Left            =   135
         TabIndex        =   17
         Top             =   2295
         Width           =   2220
      End
      Begin VB.Label Label10 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Casos prácticos"
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
         Height          =   240
         Left            =   135
         TabIndex        =   16
         Top             =   1890
         Width           =   2220
      End
      Begin VB.Label Label9 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Utilidad del contenido"
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
         Height          =   240
         Left            =   135
         TabIndex        =   15
         Top             =   1485
         Width           =   2220
      End
      Begin VB.Label Label8 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Nivel del contenido"
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
         Height          =   240
         Left            =   135
         TabIndex        =   14
         Top             =   1080
         Width           =   2220
      End
      Begin VB.Label Label7 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Organización del curso"
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
         Height          =   240
         Left            =   135
         TabIndex        =   13
         Top             =   675
         Width           =   2220
      End
      Begin VB.Label Label6 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Malo"
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
         Height          =   240
         Left            =   9135
         TabIndex        =   12
         Top             =   225
         Width           =   1275
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Indiferente"
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
         Height          =   240
         Left            =   7785
         TabIndex        =   11
         Top             =   225
         Width           =   1275
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Bueno"
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
         Height          =   240
         Left            =   6435
         TabIndex        =   10
         Top             =   225
         Width           =   1275
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Muy Bueno"
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
         Height          =   240
         Left            =   5085
         TabIndex        =   9
         Top             =   225
         Width           =   1275
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Excelente"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000C000&
         Height          =   240
         Left            =   3735
         TabIndex        =   8
         Top             =   225
         Width           =   1275
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Datos generales"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1590
      Left            =   45
      TabIndex        =   0
      Top             =   540
      Width           =   10590
      Begin VB.TextBox txtdescripcion 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   330
         Left            =   3690
         Locked          =   -1  'True
         TabIndex        =   116
         Top             =   315
         Width           =   6495
      End
      Begin VB.TextBox txtFecha 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Enabled         =   0   'False
         Height          =   330
         Left            =   1890
         Locked          =   -1  'True
         TabIndex        =   6
         Top             =   1125
         Width           =   2175
      End
      Begin VB.TextBox txtAsistente 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   330
         Left            =   1890
         Locked          =   -1  'True
         TabIndex        =   4
         Top             =   720
         Width           =   6495
      End
      Begin VB.TextBox txtCurso 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
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
         Left            =   1890
         Locked          =   -1  'True
         TabIndex        =   2
         Top             =   315
         Width           =   1770
      End
      Begin VB.Label Label1 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Fecha"
         Height          =   240
         Left            =   315
         TabIndex        =   5
         Top             =   1170
         Width           =   1320
      End
      Begin VB.Label lblAsistente 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Asistente"
         Height          =   240
         Left            =   315
         TabIndex        =   3
         Top             =   765
         Width           =   1320
      End
      Begin VB.Label lblCurso 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Curso"
         Height          =   240
         Left            =   315
         TabIndex        =   1
         Top             =   360
         Width           =   1320
      End
   End
   Begin VB.Label lblMedia 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Valoración Media:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   8055
      TabIndex        =   101
      Top             =   7875
      Visible         =   0   'False
      Width           =   1590
   End
   Begin VB.Label lblsubtitulo 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "Encuesta de satisfacción"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   300
      Left            =   225
      TabIndex        =   100
      Top             =   90
      Width           =   3045
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   9990
      Picture         =   "frmFormacion_Evaluacion.frx":6852
      Top             =   45
      Width           =   480
   End
   Begin VB.Shape fondo 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   555
      Left            =   0
      Top             =   0
      Width           =   10700
   End
   Begin VB.Label Label19 
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
      Height          =   240
      Left            =   180
      TabIndex        =   99
      Top             =   7830
      Width           =   1320
   End
   Begin VB.Line Line2 
      X1              =   9090
      X2              =   9090
      Y1              =   2205
      Y2              =   7110
   End
End
Attribute VB_Name = "frmFormacion_Evaluacion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'FORMULARIO CREADO PARA:
'MANTIS: M-0996


Public PK As Long

Public ID_ASISTENTE As Long
Private oEvaluacion As New clsFormacion_evaluacion
Private click(1 To 12) As Integer


Private Sub cmdSalir_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    log (Me.Name)
    cargar_botones Me
    
    If oEvaluacion.Carga_usuario(PK, ID_ASISTENTE) = False Then
        cmdok.Visible = True
        cmdSalir.Visible = False
        carga_formulario_alta
    Else
        cmdok.Visible = False
        cmdSalir.Visible = True
        carga_formulario
    End If

End Sub

Private Sub carga_formulario_alta()

'Reset de variable click
    For i = 1 To 12
        click(i) = 0
    Next i
    
'Carga con los valores del curso a evaluar

    Dim oCurso As New clsFormacion_cursos
    Dim oAsistente As New clsEmpleados
    
    oCurso.Carga (PK)
    oAsistente.CARGAR ID_ASISTENTE
    
    'If oCurso.getTIPO_MODALIDAD_ID = 0 Then
        txtCurso.Text = "RFI-" & Format(oCurso.getCOD_CURSO, "000") & "/" & Year(Date)
    'Else
    '    txtCurso.Text = "0301-" & Format(oCurso.getCOD_CURSO)
    'End If
    txtdescripcion.Text = oCurso.getDESCRIPCION
    txtAsistente.Text = oAsistente.getNOMBRE
    txtFecha.Text = Date
    
End Sub

Private Sub carga_formulario()

'Carga con los valores del curso evaluado

    Dim oCurso As New clsFormacion_cursos
    Dim oAsistente As New clsEmpleados
    Dim media As Integer
    Dim divisor As Integer
    
    oCurso.Carga (PK)
    oAsistente.CARGAR ID_ASISTENTE
    
    'If oCurso.getTIPO_MODALIDAD_ID = 0 Then
        txtCurso.Text = "RFI-" & Format(oCurso.getCOD_CURSO, "000") & "/" & Year(Date)
    'Else
    '    txtCurso.Text = "0301-" & Format(oCurso.getCOD_CURSO)
    'End If
    txtdescripcion.Text = oCurso.getDESCRIPCION
    
    txtAsistente.Text = oAsistente.getNOMBRE
    txtFecha.Text = Date
    divisor = 11
    
    If oEvaluacion.getORGANIZACION = 0 Then
        divisor = divisor - 1
    End If
    If oEvaluacion.getNIVEL = 0 Then
        divisor = divisor - 1
    End If
    If oEvaluacion.getUTILIDAD = 0 Then
        divisor = divisor - 1
    End If
    If oEvaluacion.getCASOS_PRACTICOS = 0 Then
        divisor = divisor - 1
    End If
    If oEvaluacion.getAUDIOVISUALES = 0 Then
        divisor = divisor - 1
    End If
    If oEvaluacion.getDINAMICAS = 0 Then
        divisor = divisor - 1
    End If
    If oEvaluacion.getCOMODIDAD = 0 Then
        divisor = divisor - 1
    End If
    If oEvaluacion.getHORARIO = 0 Then
        divisor = divisor - 1
    End If
    If oEvaluacion.getMATERIAL = 0 Then
        divisor = divisor - 1
    End If
    If oEvaluacion.getGENERAL = 0 Then
        divisor = divisor - 1
    End If
    media = oEvaluacion.getORGANIZACION + oEvaluacion.getNIVEL + oEvaluacion.getUTILIDAD + oEvaluacion.getCASOS_PRACTICOS + oEvaluacion.getAUDIOVISUALES + oEvaluacion.getDINAMICAS + oEvaluacion.getCOMODIDAD + oEvaluacion.getHORARIO + oEvaluacion.getMATERIAL + oEvaluacion.getGENERAL
    media = Round((media / divisor), 0)
    
    Select Case media
    Case 0
        txtMedia.ForeColor = &H404040        'GRIS
        txtMedia.Text = "No aplica"
    Case 1
        txtMedia.ForeColor = &H8000&     'Verde
        txtMedia.Text = "Excelente"
    Case 2
        txtMedia.ForeColor = &H80000001      'Azul
        txtMedia.Text = "Muy Bueno"
    Case 3
        txtMedia.ForeColor = &H80000001      'Azul
        txtMedia.Text = "Bueno"
    Case 4
        txtMedia.ForeColor = &H0      'Negro
        txtMedia.Text = "Indiferente"
    Case 5
        txtMedia.ForeColor = &HFF&            'Rojo
        txtMedia.Text = "Malo"
    End Select

    txtMedia.Visible = True
    lblMedia.Visible = True
    
'Lista de valores
    txtObservaciones.Text = Trim(oEvaluacion.getOBSERVACIONES)
    
    optOrganizacion.Item(oEvaluacion.getORGANIZACION).value = True
    optContenido.Item(oEvaluacion.getNIVEL).value = True
    optUtilidad.Item(oEvaluacion.getUTILIDAD).value = True
    optPracticos.Item(oEvaluacion.getCASOS_PRACTICOS).value = True
    optMedios.Item(oEvaluacion.getAUDIOVISUALES).value = True
    optGrupo.Item(oEvaluacion.getDINAMICAS).value = True
    optAula.Item(oEvaluacion.getCOMODIDAD).value = True
    optAmbiente.Item(oEvaluacion.getCOMODIDAD).value = True
    optDuracion.Item(oEvaluacion.getDURACION).value = True
    optHorario.Item(oEvaluacion.getHORARIO).value = True
    optMaterial.Item(oEvaluacion.getMATERIAL).value = True
    optGeneral.Item(oEvaluacion.getGENERAL).value = True
      
'    For i = 0 To 5
'         optOrganizacion.Item(i).Enabled = False
'         optContenido.Item(i).Enabled = False
'         optUtilidad.Item(i).Enabled = False
'         optPracticos.Item(i).Enabled = False
'         optMedios.Item(i).Enabled = False
'         optGrupo.Item(i).Enabled = False
'         optAula.Item(i).Enabled = False
'         optAmbiente.Item(i).Enabled = False
'         optDuracion.Item(i).Enabled = False
'         optHorario.Item(i).Enabled = False
'         optMaterial.Item(i).Enabled = False
'         optGeneral.Item(i).Enabled = False
'    Next i
    
    cmdok.Enabled = False
    
End Sub


Private Sub cmdcancel_Click()
    Unload Me
End Sub

Private Sub cmdOk_Click()
'Registro de la evaluación en la base de datos
    On Error GoTo cmdok_Click_Error
         
    Dim i As Integer
    Dim comprobado As Boolean
    
    comprobado = comprueba_formulario
    
    If comprobado Then
        oEvaluacion.setCURSO_ID = PK
        oEvaluacion.setEMPLEADO_ID = ID_ASISTENTE
        'M1106-I
        'oEvaluacion.setFECHA_EVALUACION = Format(Date, "yyyy-mm-dd hh:mm:ss")
        'M1106-F
        
        For i = 0 To 5
            If optOrganizacion.Item(i).value = True Then
                oEvaluacion.setORGANIZACION = i
            End If
            If optContenido.Item(i).value = True Then
                oEvaluacion.setNIVEL = i
            End If
            If optUtilidad.Item(i).value = True Then
                oEvaluacion.setUTILIDAD = i
            End If
            If optPracticos.Item(i).value = True Then
                oEvaluacion.setCASOS_PRACTICOS = i
            End If
            If optMedios.Item(i).value = True Then
                oEvaluacion.setAUDIOVISUALES = i
            End If
            If optGrupo.Item(i).value = True Then
                oEvaluacion.setDINAMICAS = i
            End If
            If optAula.Item(i).value = True Then
                oEvaluacion.setCOMODIDAD = i
            End If
            If optAula.Item(i).value = True Then
                oEvaluacion.setCOMODIDAD = i
            End If
            If optDuracion.Item(i).value = True Then
                oEvaluacion.setDURACION = i
            End If
            If optHorario.Item(i).value = True Then
                oEvaluacion.setHORARIO = i
            End If
            If optMaterial.Item(i).value = True Then
                oEvaluacion.setMATERIAL = i
            End If
            If optGeneral.Item(i).value = True Then
                oEvaluacion.setGENERAL = i
            End If
             
        Next i
        
        oEvaluacion.setOBSERVACIONES = Trim(txtObservaciones.Text)
        oEvaluacion.Insertar
        
        MsgBox "Gracias por cooperar. La evaluación se ha guardado con éxito.", vbOKOnly + vbInformation
        Unload Me
        
    Else
        MsgBox "Debe rellenar todos los campos por favor", vbExclamation, App.Title
    End If
    
    
    Exit Sub
    
cmdok_Click_Error:
  MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure cmdok_Click of Formulario frmFormacion_Evalucion Procedure cmdok_Click"
End Sub


Private Function comprueba_formulario() As Boolean
     For i = 1 To 12
        If click(i) = 0 Then
            comprueba_formulario = False
            Exit Function
        End If
     Next i
     
     comprueba_formulario = True
End Function


Private Sub Label2_Click()
    optOrganizacion(1).value = True
    optContenido(1).value = True
    optUtilidad(1).value = True
    optPracticos(1).value = True
    optMedios(1).value = True
    optGrupo(1).value = True
    optAula(1).value = True
    optAmbiente(1).value = True
    optDuracion(1).value = True
    optHorario(1).value = True
    optMaterial(1).value = True
    optGeneral(1).value = True
End Sub

Private Sub optAmbiente_Click(Index As Integer)
    click(8) = 1
End Sub

Private Sub optAula_Click(Index As Integer)
    click(7) = 1
End Sub

Private Sub optContenido_Click(Index As Integer)
    click(2) = 1
End Sub

Private Sub optDuracion_Click(Index As Integer)
    click(9) = 1
End Sub

Private Sub optGeneral_Click(Index As Integer)
    click(12) = 1
End Sub

Private Sub optGrupo_Click(Index As Integer)
    click(6) = 1
End Sub

Private Sub optHorario_Click(Index As Integer)
    click(10) = 1
End Sub

Private Sub optMaterial_Click(Index As Integer)
    click(11) = 1
End Sub

Private Sub optMedios_Click(Index As Integer)
    click(5) = 1
End Sub

Private Sub optOrganizacion_Click(Index As Integer)
    click(1) = 1
End Sub

Private Sub optPracticos_Click(Index As Integer)
    click(4) = 1
End Sub

Private Sub optUtilidad_Click(Index As Integer)
    click(3) = 1
End Sub
