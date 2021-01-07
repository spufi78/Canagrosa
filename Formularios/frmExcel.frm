VERSION 5.00
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#13.2#0"; "Codejock.Controls.v13.2.1.ocx"
Begin VB.Form frmExcel 
   Caption         =   "EXCEL"
   ClientHeight    =   11205
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   19830
   LinkTopic       =   "Form1"
   ScaleHeight     =   11205
   ScaleWidth      =   19830
   StartUpPosition =   3  'Windows Default
   Begin XtremeSuiteControls.Resizer Resizer1 
      Height          =   9600
      Left            =   3600
      TabIndex        =   0
      Top             =   270
      Width           =   15855
      _Version        =   851970
      _ExtentX        =   27966
      _ExtentY        =   16933
      _StockProps     =   1
      BorderStyle     =   1
      Begin VB.Frame Frame2 
         Caption         =   "REPETIBILIDAD DE LAS LECTURAS / CORRECCIÓN DE CALIBRACIÓN"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   4650
         Left            =   90
         TabIndex        =   1
         Top             =   180
         Width           =   13965
         Begin VB.TextBox Text1 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            Height          =   330
            Index           =   89
            Left            =   12510
            TabIndex        =   93
            Top             =   3915
            Width           =   1410
         End
         Begin VB.TextBox Text1 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            Height          =   330
            Index           =   88
            Left            =   12510
            TabIndex        =   92
            Top             =   3600
            Width           =   1410
         End
         Begin VB.TextBox Text1 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            Height          =   330
            Index           =   87
            Left            =   12510
            TabIndex        =   91
            Top             =   3285
            Width           =   1410
         End
         Begin VB.TextBox Text1 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            Height          =   330
            Index           =   86
            Left            =   12510
            TabIndex        =   90
            Top             =   2970
            Width           =   1410
         End
         Begin VB.TextBox Text1 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            Height          =   330
            Index           =   85
            Left            =   12510
            TabIndex        =   89
            Top             =   2655
            Width           =   1410
         End
         Begin VB.TextBox Text1 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            Height          =   330
            Index           =   84
            Left            =   12510
            TabIndex        =   88
            Top             =   2340
            Width           =   1410
         End
         Begin VB.TextBox Text1 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            Height          =   330
            Index           =   83
            Left            =   12510
            TabIndex        =   87
            Top             =   2025
            Width           =   1410
         End
         Begin VB.TextBox Text1 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            Height          =   330
            Index           =   82
            Left            =   12510
            TabIndex        =   86
            Top             =   1710
            Width           =   1410
         End
         Begin VB.TextBox Text1 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            Height          =   330
            Index           =   81
            Left            =   12510
            TabIndex        =   85
            Top             =   1395
            Width           =   1410
         End
         Begin VB.TextBox Text1 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            Height          =   330
            Index           =   80
            Left            =   12510
            TabIndex        =   84
            Top             =   1080
            Width           =   1410
         End
         Begin VB.TextBox Text1 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            Height          =   330
            Index           =   79
            Left            =   11115
            TabIndex        =   83
            Top             =   3915
            Width           =   1410
         End
         Begin VB.TextBox Text1 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            Height          =   330
            Index           =   78
            Left            =   11115
            TabIndex        =   82
            Top             =   3600
            Width           =   1410
         End
         Begin VB.TextBox Text1 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            Height          =   330
            Index           =   77
            Left            =   11115
            TabIndex        =   81
            Top             =   3285
            Width           =   1410
         End
         Begin VB.TextBox Text1 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            Height          =   330
            Index           =   76
            Left            =   11115
            TabIndex        =   80
            Top             =   2970
            Width           =   1410
         End
         Begin VB.TextBox Text1 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            Height          =   330
            Index           =   75
            Left            =   11115
            TabIndex        =   79
            Top             =   2655
            Width           =   1410
         End
         Begin VB.TextBox Text1 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            Height          =   330
            Index           =   74
            Left            =   11115
            TabIndex        =   78
            Top             =   2340
            Width           =   1410
         End
         Begin VB.TextBox Text1 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            Height          =   330
            Index           =   73
            Left            =   11115
            TabIndex        =   77
            Top             =   2025
            Width           =   1410
         End
         Begin VB.TextBox Text1 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            Height          =   330
            Index           =   72
            Left            =   11115
            TabIndex        =   76
            Top             =   1710
            Width           =   1410
         End
         Begin VB.TextBox Text1 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            Height          =   330
            Index           =   71
            Left            =   11115
            TabIndex        =   75
            Top             =   1395
            Width           =   1410
         End
         Begin VB.TextBox Text1 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            Height          =   330
            Index           =   70
            Left            =   11115
            TabIndex        =   74
            Top             =   1080
            Width           =   1410
         End
         Begin VB.TextBox Text1 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            Height          =   330
            Index           =   69
            Left            =   9720
            TabIndex        =   73
            Top             =   3915
            Width           =   1410
         End
         Begin VB.TextBox Text1 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            Height          =   330
            Index           =   68
            Left            =   9720
            TabIndex        =   72
            Top             =   3600
            Width           =   1410
         End
         Begin VB.TextBox Text1 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            Height          =   330
            Index           =   67
            Left            =   9720
            TabIndex        =   71
            Top             =   3285
            Width           =   1410
         End
         Begin VB.TextBox Text1 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            Height          =   330
            Index           =   66
            Left            =   9720
            TabIndex        =   70
            Top             =   2970
            Width           =   1410
         End
         Begin VB.TextBox Text1 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            Height          =   330
            Index           =   65
            Left            =   9720
            TabIndex        =   69
            Top             =   2655
            Width           =   1410
         End
         Begin VB.TextBox Text1 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            Height          =   330
            Index           =   64
            Left            =   9720
            TabIndex        =   68
            Top             =   2340
            Width           =   1410
         End
         Begin VB.TextBox Text1 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            Height          =   330
            Index           =   63
            Left            =   9720
            TabIndex        =   67
            Top             =   2025
            Width           =   1410
         End
         Begin VB.TextBox Text1 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            Height          =   330
            Index           =   62
            Left            =   9720
            TabIndex        =   66
            Top             =   1710
            Width           =   1410
         End
         Begin VB.TextBox Text1 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            Height          =   330
            Index           =   61
            Left            =   9720
            TabIndex        =   65
            Top             =   1395
            Width           =   1410
         End
         Begin VB.TextBox Text1 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            Height          =   330
            Index           =   60
            Left            =   9720
            TabIndex        =   64
            Top             =   1080
            Width           =   1410
         End
         Begin VB.TextBox Text1 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            Height          =   330
            Index           =   59
            Left            =   8280
            TabIndex        =   63
            Top             =   3915
            Width           =   1410
         End
         Begin VB.TextBox Text1 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            Height          =   330
            Index           =   58
            Left            =   8280
            TabIndex        =   62
            Top             =   3600
            Width           =   1410
         End
         Begin VB.TextBox Text1 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            Height          =   330
            Index           =   57
            Left            =   8280
            TabIndex        =   61
            Top             =   3285
            Width           =   1410
         End
         Begin VB.TextBox Text1 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            Height          =   330
            Index           =   56
            Left            =   8280
            TabIndex        =   60
            Top             =   2970
            Width           =   1410
         End
         Begin VB.TextBox Text1 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            Height          =   330
            Index           =   55
            Left            =   8280
            TabIndex        =   59
            Top             =   2655
            Width           =   1410
         End
         Begin VB.TextBox Text1 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            Height          =   330
            Index           =   54
            Left            =   8280
            TabIndex        =   58
            Top             =   2340
            Width           =   1410
         End
         Begin VB.TextBox Text1 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            Height          =   330
            Index           =   53
            Left            =   8280
            TabIndex        =   57
            Top             =   2025
            Width           =   1410
         End
         Begin VB.TextBox Text1 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            Height          =   330
            Index           =   52
            Left            =   8280
            TabIndex        =   56
            Top             =   1710
            Width           =   1410
         End
         Begin VB.TextBox Text1 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            Height          =   330
            Index           =   51
            Left            =   8280
            TabIndex        =   55
            Top             =   1395
            Width           =   1410
         End
         Begin VB.TextBox Text1 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            Height          =   330
            Index           =   50
            Left            =   8280
            TabIndex        =   54
            Top             =   1080
            Width           =   1410
         End
         Begin VB.TextBox Text1 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            Height          =   330
            Index           =   49
            Left            =   6885
            TabIndex        =   53
            Top             =   3915
            Width           =   1410
         End
         Begin VB.TextBox Text1 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            Height          =   330
            Index           =   48
            Left            =   6885
            TabIndex        =   52
            Top             =   3600
            Width           =   1410
         End
         Begin VB.TextBox Text1 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            Height          =   330
            Index           =   47
            Left            =   6885
            TabIndex        =   51
            Top             =   3285
            Width           =   1410
         End
         Begin VB.TextBox Text1 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            Height          =   330
            Index           =   46
            Left            =   6885
            TabIndex        =   50
            Top             =   2970
            Width           =   1410
         End
         Begin VB.TextBox Text1 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            Height          =   330
            Index           =   45
            Left            =   6885
            TabIndex        =   49
            Top             =   2655
            Width           =   1410
         End
         Begin VB.TextBox Text1 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            Height          =   330
            Index           =   44
            Left            =   6885
            TabIndex        =   48
            Top             =   2340
            Width           =   1410
         End
         Begin VB.TextBox Text1 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            Height          =   330
            Index           =   43
            Left            =   6885
            TabIndex        =   47
            Top             =   2025
            Width           =   1410
         End
         Begin VB.TextBox Text1 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            Height          =   330
            Index           =   42
            Left            =   6885
            TabIndex        =   46
            Top             =   1710
            Width           =   1410
         End
         Begin VB.TextBox Text1 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            Height          =   330
            Index           =   41
            Left            =   6885
            TabIndex        =   45
            Top             =   1395
            Width           =   1410
         End
         Begin VB.TextBox Text1 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            Height          =   330
            Index           =   40
            Left            =   6885
            TabIndex        =   44
            Top             =   1080
            Width           =   1410
         End
         Begin VB.TextBox Text1 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            Height          =   330
            Index           =   39
            Left            =   5490
            TabIndex        =   43
            Top             =   3915
            Width           =   1410
         End
         Begin VB.TextBox Text1 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            Height          =   330
            Index           =   38
            Left            =   5490
            TabIndex        =   42
            Top             =   3600
            Width           =   1410
         End
         Begin VB.TextBox Text1 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            Height          =   330
            Index           =   37
            Left            =   5490
            TabIndex        =   41
            Top             =   3285
            Width           =   1410
         End
         Begin VB.TextBox Text1 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            Height          =   330
            Index           =   36
            Left            =   5490
            TabIndex        =   40
            Top             =   2970
            Width           =   1410
         End
         Begin VB.TextBox Text1 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            Height          =   330
            Index           =   35
            Left            =   5490
            TabIndex        =   39
            Top             =   2655
            Width           =   1410
         End
         Begin VB.TextBox Text1 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            Height          =   330
            Index           =   34
            Left            =   5490
            TabIndex        =   38
            Top             =   2340
            Width           =   1410
         End
         Begin VB.TextBox Text1 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            Height          =   330
            Index           =   33
            Left            =   5490
            TabIndex        =   37
            Top             =   2025
            Width           =   1410
         End
         Begin VB.TextBox Text1 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            Height          =   330
            Index           =   32
            Left            =   5490
            TabIndex        =   36
            Top             =   1710
            Width           =   1410
         End
         Begin VB.TextBox Text1 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            Height          =   330
            Index           =   31
            Left            =   5490
            TabIndex        =   35
            Top             =   1395
            Width           =   1410
         End
         Begin VB.TextBox Text1 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            Height          =   330
            Index           =   30
            Left            =   5490
            TabIndex        =   34
            Top             =   1080
            Width           =   1410
         End
         Begin VB.TextBox Text1 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            Height          =   330
            Index           =   29
            Left            =   4050
            TabIndex        =   33
            Top             =   3915
            Width           =   1410
         End
         Begin VB.TextBox Text1 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            Height          =   330
            Index           =   28
            Left            =   4050
            TabIndex        =   32
            Top             =   3600
            Width           =   1410
         End
         Begin VB.TextBox Text1 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            Height          =   330
            Index           =   27
            Left            =   4050
            TabIndex        =   31
            Top             =   3285
            Width           =   1410
         End
         Begin VB.TextBox Text1 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            Height          =   330
            Index           =   26
            Left            =   4050
            TabIndex        =   30
            Top             =   2970
            Width           =   1410
         End
         Begin VB.TextBox Text1 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            Height          =   330
            Index           =   25
            Left            =   4050
            TabIndex        =   29
            Top             =   2655
            Width           =   1410
         End
         Begin VB.TextBox Text1 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            Height          =   330
            Index           =   24
            Left            =   4050
            TabIndex        =   28
            Top             =   2340
            Width           =   1410
         End
         Begin VB.TextBox Text1 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            Height          =   330
            Index           =   23
            Left            =   4050
            TabIndex        =   27
            Top             =   2025
            Width           =   1410
         End
         Begin VB.TextBox Text1 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            Height          =   330
            Index           =   22
            Left            =   4050
            TabIndex        =   26
            Top             =   1710
            Width           =   1410
         End
         Begin VB.TextBox Text1 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            Height          =   330
            Index           =   21
            Left            =   4050
            TabIndex        =   25
            Top             =   1395
            Width           =   1410
         End
         Begin VB.TextBox Text1 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            Height          =   330
            Index           =   20
            Left            =   4050
            TabIndex        =   24
            Top             =   1080
            Width           =   1410
         End
         Begin VB.TextBox Text1 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            Height          =   330
            Index           =   19
            Left            =   2655
            TabIndex        =   23
            Top             =   3915
            Width           =   1410
         End
         Begin VB.TextBox Text1 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            Height          =   330
            Index           =   18
            Left            =   2655
            TabIndex        =   22
            Top             =   3600
            Width           =   1410
         End
         Begin VB.TextBox Text1 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            Height          =   330
            Index           =   17
            Left            =   2655
            TabIndex        =   21
            Top             =   3285
            Width           =   1410
         End
         Begin VB.TextBox Text1 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            Height          =   330
            Index           =   16
            Left            =   2655
            TabIndex        =   20
            Top             =   2970
            Width           =   1410
         End
         Begin VB.TextBox Text1 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            Height          =   330
            Index           =   15
            Left            =   2655
            TabIndex        =   19
            Top             =   2655
            Width           =   1410
         End
         Begin VB.TextBox Text1 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            Height          =   330
            Index           =   14
            Left            =   2655
            TabIndex        =   18
            Top             =   2340
            Width           =   1410
         End
         Begin VB.TextBox Text1 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            Height          =   330
            Index           =   13
            Left            =   2655
            TabIndex        =   17
            Top             =   2025
            Width           =   1410
         End
         Begin VB.TextBox Text1 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            Height          =   330
            Index           =   12
            Left            =   2655
            TabIndex        =   16
            Top             =   1710
            Width           =   1410
         End
         Begin VB.TextBox Text1 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            Height          =   330
            Index           =   11
            Left            =   2655
            TabIndex        =   15
            Top             =   1395
            Width           =   1410
         End
         Begin VB.TextBox Text1 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            Height          =   330
            Index           =   10
            Left            =   2655
            TabIndex        =   14
            Top             =   1080
            Width           =   1410
         End
         Begin VB.TextBox Text1 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            Height          =   330
            Index           =   9
            Left            =   1260
            TabIndex        =   13
            Top             =   3915
            Width           =   1410
         End
         Begin VB.TextBox Text1 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            Height          =   330
            Index           =   8
            Left            =   1260
            TabIndex        =   12
            Top             =   3600
            Width           =   1410
         End
         Begin VB.TextBox Text1 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            Height          =   330
            Index           =   7
            Left            =   1260
            TabIndex        =   11
            Top             =   3285
            Width           =   1410
         End
         Begin VB.TextBox Text1 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            Height          =   330
            Index           =   6
            Left            =   1260
            TabIndex        =   10
            Top             =   2970
            Width           =   1410
         End
         Begin VB.TextBox Text1 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            Height          =   330
            Index           =   5
            Left            =   1260
            TabIndex        =   9
            Top             =   2655
            Width           =   1410
         End
         Begin VB.TextBox Text1 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            Height          =   330
            Index           =   4
            Left            =   1260
            TabIndex        =   8
            Top             =   2340
            Width           =   1410
         End
         Begin VB.TextBox Text1 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            Height          =   330
            Index           =   3
            Left            =   1260
            TabIndex        =   7
            Top             =   2025
            Width           =   1410
         End
         Begin VB.TextBox Text1 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            Height          =   330
            Index           =   2
            Left            =   1260
            TabIndex        =   6
            Top             =   1710
            Width           =   1410
         End
         Begin VB.TextBox Text1 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            Height          =   330
            Index           =   1
            Left            =   1260
            TabIndex        =   5
            Top             =   1395
            Width           =   1410
         End
         Begin VB.TextBox Text1 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            Height          =   330
            Index           =   0
            Left            =   1260
            TabIndex        =   4
            Top             =   1080
            Width           =   1410
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "M"
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
            Left            =   1755
            TabIndex        =   94
            Top             =   810
            Width           =   195
         End
      End
   End
   Begin XtremeSuiteControls.Resizer Resizer2 
      Height          =   9600
      Left            =   45
      TabIndex        =   2
      Top             =   90
      Width           =   10455
      _Version        =   851970
      _ExtentX        =   18441
      _ExtentY        =   16933
      _StockProps     =   1
      BorderStyle     =   1
      Begin VB.Frame Frame1 
         Caption         =   "CARACTERÍSTICAS DEL MENSURANDO"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3255
         Left            =   90
         TabIndex        =   3
         Top             =   90
         Width           =   10230
      End
   End
End
Attribute VB_Name = "frmExcel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

