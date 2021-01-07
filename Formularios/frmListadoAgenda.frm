VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmListadoAgenda 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Agenda"
   ClientHeight    =   8325
   ClientLeft      =   135
   ClientTop       =   1425
   ClientWidth     =   11775
   Icon            =   "frmListadoAgenda.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8325
   ScaleWidth      =   11775
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   "Filtro"
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
      Height          =   690
      Left            =   45
      TabIndex        =   33
      Top             =   360
      Width           =   11625
      Begin VB.TextBox txtnombre 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   765
         TabIndex        =   34
         Top             =   225
         Width           =   10680
      End
      Begin VB.Label Label2 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Nombre"
         Height          =   195
         Index           =   0
         Left            =   90
         TabIndex        =   35
         Top             =   270
         Width           =   615
      End
   End
   Begin VB.CommandButton cmdanadir 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Añadir"
      Height          =   885
      Left            =   90
      Style           =   1  'Graphical
      TabIndex        =   29
      Top             =   7410
      Width           =   1155
   End
   Begin VB.CommandButton cmdcancel 
      BackColor       =   &H00E0E0E0&
      Caption         =   "ESC-Salir"
      Height          =   885
      Left            =   10560
      Style           =   1  'Graphical
      TabIndex        =   32
      Top             =   7410
      Width           =   1155
   End
   Begin VB.CommandButton cmdmodificar 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Modificar"
      Height          =   885
      Left            =   1290
      Style           =   1  'Graphical
      TabIndex        =   30
      Top             =   7410
      Width           =   1155
   End
   Begin VB.CommandButton cmdeliminar 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Eliminar"
      Height          =   885
      Left            =   2490
      Style           =   1  'Graphical
      TabIndex        =   31
      Top             =   7410
      Width           =   1155
   End
   Begin MSComctlLib.ListView lista 
      Height          =   6225
      Left            =   45
      TabIndex        =   28
      Top             =   1095
      Width           =   11640
      _ExtentX        =   20532
      _ExtentY        =   10980
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
   Begin VB.Label lbltitulo 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0FF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Agenda"
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
      Height          =   315
      Left            =   60
      TabIndex        =   0
      Top             =   30
      Width           =   11655
   End
   Begin VB.Label y 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      Caption         =   "Y"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   10800
      TabIndex        =   26
      Top             =   540
      Visible         =   0   'False
      Width           =   210
   End
   Begin VB.Label z 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      Caption         =   "Z"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   11220
      TabIndex        =   27
      Top             =   540
      Visible         =   0   'False
      Width           =   210
   End
   Begin VB.Label v 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      Caption         =   "V"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   9540
      TabIndex        =   23
      Top             =   540
      Visible         =   0   'False
      Width           =   225
   End
   Begin VB.Label w 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      Caption         =   "W"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   9900
      TabIndex        =   24
      Top             =   540
      Visible         =   0   'False
      Width           =   300
   End
   Begin VB.Label u 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      Caption         =   "U"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   9120
      TabIndex        =   22
      Top             =   540
      Visible         =   0   'False
      Width           =   225
   End
   Begin VB.Label t 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      Caption         =   "T"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   8700
      TabIndex        =   21
      Top             =   540
      Visible         =   0   'False
      Width           =   210
   End
   Begin VB.Label o 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      Caption         =   "O"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   6600
      TabIndex        =   16
      Top             =   540
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Label ñ 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      Caption         =   "Ñ"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   6180
      TabIndex        =   15
      Top             =   540
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Label p 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      Caption         =   "P"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   7020
      TabIndex        =   17
      Top             =   540
      Visible         =   0   'False
      Width           =   210
   End
   Begin VB.Label q 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      Caption         =   "Q"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   7440
      TabIndex        =   18
      Top             =   540
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Label r 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      Caption         =   "R"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   7860
      TabIndex        =   19
      Top             =   540
      Visible         =   0   'False
      Width           =   225
   End
   Begin VB.Label x 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      Caption         =   "X"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   10380
      TabIndex        =   25
      Top             =   540
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Label s 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      Caption         =   "S"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   8280
      TabIndex        =   20
      Top             =   540
      Visible         =   0   'False
      Width           =   210
   End
   Begin VB.Line Line1 
      Index           =   23
      Visible         =   0   'False
      X1              =   9420
      X2              =   9420
      Y1              =   420
      Y2              =   960
   End
   Begin VB.Line Line1 
      Index           =   22
      Visible         =   0   'False
      X1              =   11100
      X2              =   11100
      Y1              =   420
      Y2              =   960
   End
   Begin VB.Line Line1 
      Index           =   21
      Visible         =   0   'False
      X1              =   10680
      X2              =   10680
      Y1              =   420
      Y2              =   960
   End
   Begin VB.Line Line1 
      Index           =   20
      Visible         =   0   'False
      X1              =   10260
      X2              =   10260
      Y1              =   420
      Y2              =   960
   End
   Begin VB.Line Line1 
      Index           =   19
      Visible         =   0   'False
      X1              =   8160
      X2              =   8160
      Y1              =   420
      Y2              =   960
   End
   Begin VB.Line Line1 
      Index           =   18
      Visible         =   0   'False
      X1              =   7320
      X2              =   7320
      Y1              =   420
      Y2              =   960
   End
   Begin VB.Line Line1 
      Index           =   17
      Visible         =   0   'False
      X1              =   6480
      X2              =   6480
      Y1              =   420
      Y2              =   960
   End
   Begin VB.Line Line1 
      Index           =   16
      Visible         =   0   'False
      X1              =   6900
      X2              =   6900
      Y1              =   420
      Y2              =   960
   End
   Begin VB.Line Line1 
      Index           =   15
      Visible         =   0   'False
      X1              =   9000
      X2              =   9000
      Y1              =   420
      Y2              =   960
   End
   Begin VB.Line Line1 
      Index           =   14
      Visible         =   0   'False
      X1              =   8580
      X2              =   8580
      Y1              =   420
      Y2              =   960
   End
   Begin VB.Line Line1 
      Index           =   13
      Visible         =   0   'False
      X1              =   9840
      X2              =   9840
      Y1              =   420
      Y2              =   960
   End
   Begin VB.Line Line1 
      Index           =   12
      Visible         =   0   'False
      X1              =   1020
      X2              =   1020
      Y1              =   420
      Y2              =   960
   End
   Begin VB.Line Line1 
      Index           =   11
      Visible         =   0   'False
      X1              =   1440
      X2              =   1440
      Y1              =   420
      Y2              =   960
   End
   Begin VB.Line Line1 
      Index           =   10
      Visible         =   0   'False
      X1              =   1860
      X2              =   1860
      Y1              =   420
      Y2              =   960
   End
   Begin VB.Line Line1 
      Index           =   9
      Visible         =   0   'False
      X1              =   2280
      X2              =   2280
      Y1              =   420
      Y2              =   960
   End
   Begin VB.Line Line1 
      Index           =   8
      Visible         =   0   'False
      X1              =   2700
      X2              =   2700
      Y1              =   420
      Y2              =   960
   End
   Begin VB.Line Line1 
      Index           =   7
      Visible         =   0   'False
      X1              =   3120
      X2              =   3120
      Y1              =   420
      Y2              =   960
   End
   Begin VB.Line Line1 
      Index           =   6
      Visible         =   0   'False
      X1              =   3540
      X2              =   3540
      Y1              =   420
      Y2              =   960
   End
   Begin VB.Line Line1 
      Index           =   5
      Visible         =   0   'False
      X1              =   3960
      X2              =   3960
      Y1              =   420
      Y2              =   960
   End
   Begin VB.Line Line1 
      Index           =   4
      Visible         =   0   'False
      X1              =   4380
      X2              =   4380
      Y1              =   420
      Y2              =   960
   End
   Begin VB.Line Line1 
      Index           =   3
      Visible         =   0   'False
      X1              =   4800
      X2              =   4800
      Y1              =   420
      Y2              =   960
   End
   Begin VB.Line Line1 
      Index           =   2
      Visible         =   0   'False
      X1              =   5220
      X2              =   5220
      Y1              =   420
      Y2              =   960
   End
   Begin VB.Line Line1 
      Index           =   1
      Visible         =   0   'False
      X1              =   5640
      X2              =   5640
      Y1              =   420
      Y2              =   960
   End
   Begin VB.Line Line3 
      Visible         =   0   'False
      X1              =   6060
      X2              =   6060
      Y1              =   420
      Y2              =   960
   End
   Begin VB.Line Line2 
      Visible         =   0   'False
      X1              =   7740
      X2              =   7740
      Y1              =   420
      Y2              =   960
   End
   Begin VB.Label b 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      Caption         =   "B"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   720
      TabIndex        =   2
      Top             =   540
      Visible         =   0   'False
      Width           =   210
   End
   Begin VB.Label c 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      Caption         =   "C"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   1140
      TabIndex        =   3
      Top             =   540
      Visible         =   0   'False
      Width           =   225
   End
   Begin VB.Label d 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      Caption         =   "D"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   1560
      TabIndex        =   4
      Top             =   540
      Visible         =   0   'False
      Width           =   225
   End
   Begin VB.Label e 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      Caption         =   "E"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   1980
      TabIndex        =   5
      Top             =   540
      Visible         =   0   'False
      Width           =   225
   End
   Begin VB.Label f 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      Caption         =   "F"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   2400
      TabIndex        =   6
      Top             =   540
      Visible         =   0   'False
      Width           =   210
   End
   Begin VB.Label g 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      Caption         =   "G"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   2820
      TabIndex        =   7
      Top             =   540
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Label h 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      Caption         =   "H"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   3240
      TabIndex        =   8
      Top             =   540
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Label i 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      Caption         =   "I"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   3720
      TabIndex        =   9
      Top             =   540
      Visible         =   0   'False
      Width           =   90
   End
   Begin VB.Label n 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      Caption         =   "N"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   5760
      TabIndex        =   14
      Top             =   540
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Label m 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      Caption         =   "M"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   5340
      TabIndex        =   13
      Top             =   540
      Visible         =   0   'False
      Width           =   270
   End
   Begin VB.Label j 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      Caption         =   "J"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   4080
      TabIndex        =   10
      Top             =   540
      Visible         =   0   'False
      Width           =   165
   End
   Begin VB.Label k 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      Caption         =   "K"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   4500
      TabIndex        =   11
      Top             =   540
      Visible         =   0   'False
      Width           =   210
   End
   Begin VB.Label l 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      Caption         =   "L"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   4920
      TabIndex        =   12
      Top             =   540
      Visible         =   0   'False
      Width           =   180
   End
   Begin VB.Line Line1 
      Index           =   0
      Visible         =   0   'False
      X1              =   600
      X2              =   600
      Y1              =   420
      Y2              =   960
   End
   Begin VB.Label a 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      Caption         =   "A"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   300
      TabIndex        =   1
      Top             =   540
      Visible         =   0   'False
      Width           =   225
   End
   Begin VB.Shape Shape1 
      Height          =   555
      Left            =   180
      Top             =   420
      Visible         =   0   'False
      Width           =   11355
   End
End
Attribute VB_Name = "frmListadoAgenda"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdAnadir_Click()
    gAgenda = 0
    frmAgenda.Show 1
    If lista.ListItems.Count > 0 Then
'       buscar_agenda (Asc(Left(lista.ListItems(lista.selectedItem.Index), 1)))
       buscar_agenda
    End If
End Sub

Private Sub cmdcancel_Click()
    Unload Me
End Sub

Private Sub cmdEliminar_Click()
    If lista.ListItems.Count = 0 Then
        Exit Sub
    End If
    If MsgBox("Va a ELIMINAR de la agenda " & lista.ListItems(lista.selectedItem.Index) & ". ¿Esta seguro?", vbExclamation + vbYesNo, "Informacion") = vbYes Then
        Dim oAgenda As New clsAgenda
        oAgenda.setID_AGENDA = CInt(lista.ListItems(lista.selectedItem.Index).SubItems(4))
        If oAgenda.Eliminar = True Then
            If lista.ListItems.Count > 0 Then
'                buscar_agenda (Asc(Left(lista.ListItems(lista.selectedItem.Index), 1)))
                buscar_agenda
            Else
                lista.ListItems.Clear
            End If
        End If
        Set oAgenda = Nothing
    End If
    lista.SetFocus
End Sub

Private Sub cmdModificar_Click()
    If lista.ListItems.Count = 0 Then
        Exit Sub
    End If
    gAgenda = lista.ListItems(lista.selectedItem.Index).SubItems(4)
    frmAgenda.Show 1
    If lista.ListItems.Count > 0 Then
'       buscar_agenda (Asc(Left(lista.ListItems(lista.selectedItem.Index), 1)))
       buscar_agenda
    End If
End Sub

Private Sub Form_Activate()
'    a.BackColor = vbWhite
    cabecera
'    buscar_agenda (65)
    buscar_agenda
    Me.SetFocus
End Sub

'Private Sub borrar_seleccionado()
'    a.BackColor = &HC0C0C0
'    b.BackColor = &HC0C0C0
'    c.BackColor = &HC0C0C0
'    d.BackColor = &HC0C0C0
'    e.BackColor = &HC0C0C0
'    f.BackColor = &HC0C0C0
'    g.BackColor = &HC0C0C0
'    h.BackColor = &HC0C0C0
'    i.BackColor = &HC0C0C0
'    j.BackColor = &HC0C0C0
'    k.BackColor = &HC0C0C0
'    l.BackColor = &HC0C0C0
'    m.BackColor = &HC0C0C0
'    n.BackColor = &HC0C0C0
'    ñ.BackColor = &HC0C0C0
'    o.BackColor = &HC0C0C0
'    p.BackColor = &HC0C0C0
'    q.BackColor = &HC0C0C0
'    r.BackColor = &HC0C0C0
'    s.BackColor = &HC0C0C0
'    t.BackColor = &HC0C0C0
'    u.BackColor = &HC0C0C0
'    v.BackColor = &HC0C0C0
'    w.BackColor = &HC0C0C0
'    x.BackColor = &HC0C0C0
'    y.BackColor = &HC0C0C0
'    z.BackColor = &HC0C0C0
'End Sub

'Private Sub buscar_agenda(letra As String)
Private Sub buscar_agenda()
    Dim oAgenda As New clsAgenda
    Dim rs As ADODB.Recordset
'    Set rs = oAgenda.Listado_por_letra(UCase(Chr(letra)))
    Set rs = oAgenda.Listado_por_letra(txtnombre)
    lista.ListItems.Clear
    If rs.RecordCount > 0 Then
        Do
           With lista.ListItems.Add(, , rs(0))
            .SubItems(1) = rs(1)
            .SubItems(2) = rs(2)
            .SubItems(3) = rs(3)
            .SubItems(4) = rs(4)
           End With
           rs.MoveNext
        Loop Until rs.EOF
    End If
End Sub

'Private Sub Form_KeyPress(KeyAscii As Integer)
'    borrar_seleccionado
'    Select Case Chr(KeyAscii)
'    Case "a", "A"
'        a.BackColor = &HFFFFFF
'    Case "b", "B"
'        b.BackColor = &HFFFFFF
'    Case "c", "C"
'        c.BackColor = &HFFFFFF
'    Case "d", "D"
'        d.BackColor = &HFFFFFF
'    Case "e", "E"
'        e.BackColor = &HFFFFFF
'    Case "f", "F"
'        f.BackColor = &HFFFFFF
'    Case "g", "G"
'        g.BackColor = &HFFFFFF
'    Case "h", "H"
'        h.BackColor = &HFFFFFF
'    Case "i", "I"
'        i.BackColor = &HFFFFFF
'    Case "j", "J"
'        j.BackColor = &HFFFFFF
'    Case "k", "K"
'        k.BackColor = &HFFFFFF
'    Case "l", "L"
'        l.BackColor = &HFFFFFF
'    Case "m", "M"
'        m.BackColor = &HFFFFFF
'    Case "n", "N"
'        n.BackColor = &HFFFFFF
'    Case "ñ", "Ñ"
'        ñ.BackColor = &HFFFFFF
'    Case "o", "O"
'        o.BackColor = &HFFFFFF
'    Case "p", "P"
'        p.BackColor = &HFFFFFF
'    Case "q", "Q"
'        q.BackColor = &HFFFFFF
'    Case "r", "R"
'        r.BackColor = &HFFFFFF
'    Case "s", "S"
'        s.BackColor = &HFFFFFF
'    Case "t", "T"
'        t.BackColor = &HFFFFFF
'    Case "u", "U"
'        u.BackColor = &HFFFFFF
'    Case "v", "V"
'        v.BackColor = &HFFFFFF
'    Case "w", "W"
'        w.BackColor = &HFFFFFF
'    Case "x", "X"
'        x.BackColor = &HFFFFFF
'    Case "y", "Y"
'        y.BackColor = &HFFFFFF
'    Case "z", "Z"
'        z.BackColor = &HFFFFFF
'    End Select
'    buscar_agenda (KeyAscii)
'    KeyAscii = 0
'End Sub

Private Sub Form_Load()
    log (Me.Name)
    cargar_botones Me
    Me.Left = 50
    Me.Top = 50
End Sub
Private Sub cabecera()
    With lista.ColumnHeaders.Add(, , "Nombre", 5300, lvwColumnLeft)
        .Tag = "Nombre"
    End With
    With lista.ColumnHeaders.Add(, , "Teléfono", 2000, lvwColumnCenter)
        .Tag = "Teléfono"
    End With
    With lista.ColumnHeaders.Add(, , "Móvil", 2000, lvwColumnCenter)
        .Tag = "Móvil"
    End With
    With lista.ColumnHeaders.Add(, , "Fax", 2000, lvwColumnCenter)
        .Tag = "Fax"
    End With
    With lista.ColumnHeaders.Add(, , "ID", 1, lvwColumnCenter)
        .Tag = "ID"
    End With
End Sub

Private Sub txtnombre_Change()
    buscar_agenda
End Sub
