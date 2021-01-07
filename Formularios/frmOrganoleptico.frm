VERSION 5.00
Begin VB.Form frmOrganoleptico 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Organoleptico"
   ClientHeight    =   8955
   ClientLeft      =   1740
   ClientTop       =   1755
   ClientWidth     =   12570
   Icon            =   "frmOrganoleptico.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8955
   ScaleWidth      =   12570
   Begin VB.CommandButton cmdObservador 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Observador"
      Height          =   870
      Left            =   90
      Picture         =   "frmOrganoleptico.frx":1272
      Style           =   1  'Graphical
      TabIndex        =   128
      Tag             =   "Añade campo o modifica el campo existente con el mismo nombre"
      Top             =   8070
      Width           =   1095
   End
   Begin VB.CommandButton cmdok 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Aceptar"
      Height          =   870
      Left            =   10290
      Style           =   1  'Graphical
      TabIndex        =   127
      Top             =   8070
      Width           =   1050
   End
   Begin VB.CommandButton cmdcancel 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Cancelar"
      Height          =   870
      Left            =   11415
      Style           =   1  'Graphical
      TabIndex        =   126
      Top             =   8070
      Width           =   1050
   End
   Begin VB.Frame Frame7 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      ForeColor       =   &H80000008&
      Height          =   675
      Left            =   60
      TabIndex        =   123
      Top             =   7380
      Width           =   12405
      Begin VB.CommandButton cmdCalcular 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Caption         =   "Calcular"
         Height          =   375
         Left            =   10740
         TabIndex        =   22
         Top             =   210
         Width           =   1545
      End
      Begin VB.TextBox por 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Index           =   19
         Left            =   8550
         TabIndex        =   46
         Top             =   180
         Width           =   1275
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "%"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   6
         Left            =   9990
         TabIndex        =   125
         Top             =   300
         Width           =   225
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "ELEMENTOS QUE NO SON CEREALES DE BASE DE CALIDAD IRREPROCHABLE"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   6
         Left            =   180
         TabIndex        =   124
         Top             =   270
         Width           =   7095
      End
   End
   Begin VB.Frame Frame6 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      ForeColor       =   &H80000008&
      Height          =   1305
      Left            =   6300
      TabIndex        =   121
      Top             =   6060
      Width           =   6165
      Begin VB.TextBox val 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1035
         Index           =   20
         Left            =   90
         MultiLine       =   -1  'True
         TabIndex        =   21
         Top             =   180
         Width           =   5985
      End
   End
   Begin VB.TextBox val 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Index           =   0
      Left            =   3180
      TabIndex        =   0
      Top             =   390
      Width           =   1065
   End
   Begin VB.Frame Frame5 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      ForeColor       =   &H80000008&
      Height          =   1485
      Left            =   6300
      TabIndex        =   108
      Top             =   4260
      Width           =   6165
      Begin VB.TextBox por 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Index           =   16
         Left            =   4680
         TabIndex        =   43
         Top             =   990
         Width           =   945
      End
      Begin VB.TextBox val 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Index           =   16
         Left            =   3150
         TabIndex        =   19
         Top             =   990
         Width           =   1065
      End
      Begin VB.TextBox por 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Index           =   15
         Left            =   4680
         TabIndex        =   42
         Top             =   570
         Width           =   945
      End
      Begin VB.TextBox val 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Index           =   15
         Left            =   3150
         TabIndex        =   18
         Top             =   570
         Width           =   1065
      End
      Begin VB.TextBox por 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Index           =   14
         Left            =   4680
         TabIndex        =   41
         Top             =   150
         Width           =   945
      End
      Begin VB.TextBox val 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Index           =   14
         Left            =   3150
         TabIndex        =   17
         Top             =   150
         Width           =   1065
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Granos harinosos"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   116
         Left            =   180
         TabIndex        =   120
         Top             =   1080
         Width           =   1500
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "%"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   23
         Left            =   5730
         TabIndex        =   113
         Top             =   1110
         Width           =   225
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "%"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   22
         Left            =   5730
         TabIndex        =   112
         Top             =   690
         Width           =   225
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Granos berrendos"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   115
         Left            =   180
         TabIndex        =   111
         Top             =   660
         Width           =   1545
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "%"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   21
         Left            =   5730
         TabIndex        =   110
         Top             =   270
         Width           =   225
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Granos vitreos"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   114
         Left            =   180
         TabIndex        =   109
         Top             =   240
         Width           =   1260
      End
   End
   Begin VB.Frame Frame4 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      ForeColor       =   &H80000008&
      Height          =   3195
      Left            =   6300
      TabIndex        =   88
      Top             =   720
      Width           =   6165
      Begin VB.TextBox val 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Index           =   11
         Left            =   3150
         TabIndex        =   15
         Top             =   2280
         Width           =   1065
      End
      Begin VB.TextBox por 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Index           =   11
         Left            =   4680
         TabIndex        =   35
         Top             =   2280
         Width           =   945
      End
      Begin VB.TextBox val 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Index           =   6
         Left            =   3150
         TabIndex        =   10
         Top             =   180
         Width           =   1065
      End
      Begin VB.TextBox por 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Index           =   6
         Left            =   4680
         TabIndex        =   30
         Top             =   180
         Width           =   945
      End
      Begin VB.TextBox val 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Index           =   7
         Left            =   3150
         TabIndex        =   11
         Top             =   600
         Width           =   1065
      End
      Begin VB.TextBox por 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Index           =   7
         Left            =   4680
         TabIndex        =   31
         Top             =   600
         Width           =   945
      End
      Begin VB.TextBox val 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Index           =   8
         Left            =   3150
         TabIndex        =   12
         Top             =   1020
         Width           =   1065
      End
      Begin VB.TextBox por 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Index           =   8
         Left            =   4680
         TabIndex        =   32
         Top             =   1020
         Width           =   945
      End
      Begin VB.TextBox val 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Index           =   9
         Left            =   3150
         TabIndex        =   13
         Top             =   1440
         Width           =   1065
      End
      Begin VB.TextBox por 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Index           =   9
         Left            =   4680
         TabIndex        =   33
         Top             =   1440
         Width           =   945
      End
      Begin VB.TextBox val 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Index           =   10
         Left            =   3150
         TabIndex        =   14
         Top             =   1860
         Width           =   1065
      End
      Begin VB.TextBox por 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Index           =   10
         Left            =   4680
         TabIndex        =   34
         Top             =   1860
         Width           =   945
      End
      Begin VB.TextBox total 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Index           =   2
         Left            =   3150
         TabIndex        =   16
         Top             =   2700
         Width           =   1065
      End
      Begin VB.TextBox total 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Index           =   3
         Left            =   4680
         TabIndex        =   36
         Top             =   2700
         Width           =   945
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Insectos muertos y fragmentos"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   111
         Left            =   180
         TabIndex        =   117
         Top             =   2370
         Width           =   2685
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "gr"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   24
         Left            =   4320
         TabIndex        =   116
         Top             =   2430
         Width           =   210
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "%"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   24
         Left            =   5730
         TabIndex        =   115
         Top             =   2400
         Width           =   225
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Sem. de malas hierbas nocivas"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   106
         Left            =   180
         TabIndex        =   106
         Top             =   270
         Width           =   2685
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "gr"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   20
         Left            =   4320
         TabIndex        =   105
         Top             =   330
         Width           =   210
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "%"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   20
         Left            =   5730
         TabIndex        =   104
         Top             =   300
         Width           =   225
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Sem. de malas hierbas no nocivas"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   107
         Left            =   180
         TabIndex        =   103
         Top             =   690
         Width           =   2955
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "gr"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   19
         Left            =   4320
         TabIndex        =   102
         Top             =   750
         Width           =   210
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "%"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   19
         Left            =   5730
         TabIndex        =   101
         Top             =   720
         Width           =   225
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Averiados por secado"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   108
         Left            =   180
         TabIndex        =   100
         Top             =   1110
         Width           =   1860
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "gr"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   18
         Left            =   4320
         TabIndex        =   99
         Top             =   1170
         Width           =   210
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "%"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   18
         Left            =   5730
         TabIndex        =   98
         Top             =   1140
         Width           =   225
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Otros granos dañados"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   109
         Left            =   180
         TabIndex        =   97
         Top             =   1530
         Width           =   1890
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "gr"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   15
         Left            =   4320
         TabIndex        =   96
         Top             =   1590
         Width           =   210
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "%"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   15
         Left            =   5730
         TabIndex        =   95
         Top             =   1560
         Width           =   225
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Impurezas p.m.d."
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   110
         Left            =   180
         TabIndex        =   94
         Top             =   1950
         Width           =   1530
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "gr"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   14
         Left            =   4320
         TabIndex        =   93
         Top             =   2010
         Width           =   210
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "%"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   14
         Left            =   5730
         TabIndex        =   92
         Top             =   1980
         Width           =   225
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "TOTAL"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   195
         Index           =   13
         Left            =   2160
         TabIndex        =   91
         Top             =   2790
         Width           =   615
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "gr"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   13
         Left            =   4320
         TabIndex        =   90
         Top             =   2850
         Width           =   210
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "%"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   13
         Left            =   5730
         TabIndex        =   89
         Top             =   2820
         Width           =   225
      End
   End
   Begin VB.Frame Frame3 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      ForeColor       =   &H80000008&
      Height          =   1485
      Left            =   60
      TabIndex        =   68
      Top             =   4260
      Width           =   6165
      Begin VB.TextBox val 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Index           =   12
         Left            =   3150
         TabIndex        =   6
         Top             =   150
         Width           =   1065
      End
      Begin VB.TextBox por 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Index           =   12
         Left            =   4680
         TabIndex        =   37
         Top             =   150
         Width           =   945
      End
      Begin VB.TextBox val 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Index           =   13
         Left            =   3150
         TabIndex        =   7
         Top             =   570
         Width           =   1065
      End
      Begin VB.TextBox por 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Index           =   13
         Left            =   4680
         TabIndex        =   38
         Top             =   570
         Width           =   945
      End
      Begin VB.TextBox total 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Index           =   4
         Left            =   3150
         TabIndex        =   39
         Top             =   990
         Width           =   1065
      End
      Begin VB.TextBox total 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Index           =   5
         Left            =   4680
         TabIndex        =   40
         Top             =   990
         Width           =   945
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Granos maculados"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   112
         Left            =   180
         TabIndex        =   77
         Top             =   240
         Width           =   1590
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "gr"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   17
         Left            =   4320
         TabIndex        =   76
         Top             =   300
         Width           =   210
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "%"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   17
         Left            =   5730
         TabIndex        =   75
         Top             =   270
         Width           =   225
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Granos fusariados"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   113
         Left            =   180
         TabIndex        =   74
         Top             =   660
         Width           =   1560
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "gr"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   16
         Left            =   4320
         TabIndex        =   73
         Top             =   720
         Width           =   210
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "%"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   16
         Left            =   5730
         TabIndex        =   72
         Top             =   690
         Width           =   225
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "TOTAL"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   195
         Index           =   12
         Left            =   2160
         TabIndex        =   71
         Top             =   1110
         Width           =   615
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "gr"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   12
         Left            =   4320
         TabIndex        =   70
         Top             =   1140
         Width           =   210
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "%"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   12
         Left            =   5730
         TabIndex        =   69
         Top             =   1110
         Width           =   225
      End
   End
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      ForeColor       =   &H80000008&
      Height          =   2775
      Left            =   60
      TabIndex        =   49
      Top             =   1140
      Width           =   6165
      Begin VB.TextBox total 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Index           =   1
         Left            =   4680
         TabIndex        =   29
         Top             =   2280
         Width           =   945
      End
      Begin VB.TextBox total 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Index           =   0
         Left            =   3150
         TabIndex        =   28
         Top             =   2280
         Width           =   1065
      End
      Begin VB.TextBox por 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Index           =   5
         Left            =   4680
         TabIndex        =   27
         Top             =   1860
         Width           =   945
      End
      Begin VB.TextBox val 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Index           =   5
         Left            =   3150
         TabIndex        =   5
         Top             =   1860
         Width           =   1065
      End
      Begin VB.TextBox por 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Index           =   4
         Left            =   4680
         TabIndex        =   26
         Top             =   1440
         Width           =   945
      End
      Begin VB.TextBox val 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Index           =   4
         Left            =   3150
         TabIndex        =   4
         Top             =   1440
         Width           =   1065
      End
      Begin VB.TextBox por 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Index           =   3
         Left            =   4680
         TabIndex        =   25
         Top             =   1020
         Width           =   945
      End
      Begin VB.TextBox val 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Index           =   3
         Left            =   3150
         TabIndex        =   3
         Top             =   1020
         Width           =   1065
      End
      Begin VB.TextBox por 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Index           =   2
         Left            =   4680
         TabIndex        =   24
         Top             =   600
         Width           =   945
      End
      Begin VB.TextBox val 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Index           =   2
         Left            =   3150
         TabIndex        =   2
         Top             =   600
         Width           =   1065
      End
      Begin VB.TextBox por 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Index           =   1
         Left            =   4680
         TabIndex        =   23
         Top             =   180
         Width           =   945
      End
      Begin VB.TextBox val 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Index           =   1
         Left            =   3150
         TabIndex        =   1
         Top             =   180
         Width           =   1065
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "%"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   5
         Left            =   5730
         TabIndex        =   67
         Top             =   2400
         Width           =   225
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "gr"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   5
         Left            =   4320
         TabIndex        =   66
         Top             =   2430
         Width           =   210
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "TOTAL"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   195
         Index           =   5
         Left            =   2160
         TabIndex        =   65
         Top             =   2370
         Width           =   615
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "%"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   4
         Left            =   5730
         TabIndex        =   64
         Top             =   1980
         Width           =   225
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "gr"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   4
         Left            =   4320
         TabIndex        =   63
         Top             =   2010
         Width           =   210
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Con coloración del germen"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   105
         Left            =   180
         TabIndex        =   62
         Top             =   1950
         Width           =   2310
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "%"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   3
         Left            =   5730
         TabIndex        =   61
         Top             =   1560
         Width           =   225
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "gr"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   3
         Left            =   4320
         TabIndex        =   60
         Top             =   1590
         Width           =   210
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Atacados por depredadores"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   104
         Left            =   180
         TabIndex        =   59
         Top             =   1530
         Width           =   2370
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "%"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   2
         Left            =   5730
         TabIndex        =   58
         Top             =   1140
         Width           =   225
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "gr"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   2
         Left            =   4320
         TabIndex        =   57
         Top             =   1170
         Width           =   210
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Trigo blando"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   103
         Left            =   180
         TabIndex        =   56
         Top             =   1110
         Width           =   1065
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "%"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   1
         Left            =   5730
         TabIndex        =   55
         Top             =   720
         Width           =   225
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "gr"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   1
         Left            =   4320
         TabIndex        =   54
         Top             =   750
         Width           =   210
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Otros cereales"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   102
         Left            =   180
         TabIndex        =   53
         Top             =   690
         Width           =   1245
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "%"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   0
         Left            =   5730
         TabIndex        =   52
         Top             =   300
         Width           =   225
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "gr"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   0
         Left            =   4320
         TabIndex        =   51
         Top             =   330
         Width           =   210
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Granos mermados"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   101
         Left            =   180
         TabIndex        =   50
         Top             =   270
         Width           =   1590
      End
   End
   Begin VB.Frame Frame2 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      ForeColor       =   &H80000008&
      Height          =   1575
      Left            =   60
      TabIndex        =   79
      Top             =   5790
      Width           =   6165
      Begin VB.TextBox val 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Index           =   17
         Left            =   3150
         TabIndex        =   8
         Top             =   180
         Width           =   1065
      End
      Begin VB.TextBox por 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Index           =   17
         Left            =   4680
         TabIndex        =   44
         Top             =   180
         Width           =   945
      End
      Begin VB.TextBox val 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Index           =   18
         Left            =   3150
         TabIndex        =   9
         Top             =   600
         Width           =   1065
      End
      Begin VB.TextBox por 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Index           =   18
         Left            =   4680
         TabIndex        =   45
         Top             =   600
         Width           =   945
      End
      Begin VB.TextBox val 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Index           =   19
         Left            =   4680
         TabIndex        =   20
         Top             =   1020
         Width           =   945
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "GRANOS PARTIDOS"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   117
         Left            =   180
         TabIndex        =   87
         Top             =   270
         Width           =   1725
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "gr"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   11
         Left            =   4320
         TabIndex        =   86
         Top             =   330
         Width           =   210
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "%"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   11
         Left            =   5730
         TabIndex        =   85
         Top             =   300
         Width           =   225
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "GRANOS GERMINADOS"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   118
         Left            =   180
         TabIndex        =   84
         Top             =   690
         Width           =   2010
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "gr"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   10
         Left            =   4320
         TabIndex        =   83
         Top             =   750
         Width           =   210
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "%"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   10
         Left            =   5730
         TabIndex        =   82
         Top             =   720
         Width           =   225
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "PESO DE LOS 1000 GRANOS"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   119
         Left            =   180
         TabIndex        =   81
         Top             =   1110
         Width           =   2460
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "gr"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   9
         Left            =   5730
         TabIndex        =   80
         Top             =   1170
         Width           =   210
      End
   End
   Begin VB.Label lblCerrada 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
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
      Height          =   300
      Left            =   9660
      TabIndex        =   129
      Top             =   45
      Width           =   2805
   End
   Begin VB.Label lbldeter 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Caption         =   "OBSERVACIONES"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   300
      Index           =   1
      Left            =   6300
      TabIndex        =   122
      Top             =   5760
      Width           =   6165
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      Caption         =   "gr"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Index           =   100
      Left            =   4380
      TabIndex        =   119
      Top             =   480
      Width           =   210
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0C0C0&
      Caption         =   "PESO DE MUESTRA"
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
      Index           =   25
      Left            =   60
      TabIndex        =   118
      Top             =   450
      Width           =   2790
   End
   Begin VB.Label lbldeter 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Caption         =   "GRANOS MACULADOS Y/O FUSARIADOS"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   300
      Index           =   2
      Left            =   60
      TabIndex        =   78
      Top             =   3930
      Width           =   6165
   End
   Begin VB.Label lbldeter 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Caption         =   "EXAMEN AL CORTAGRANO"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   300
      Index           =   4
      Left            =   6300
      TabIndex        =   114
      Top             =   3930
      Width           =   6165
   End
   Begin VB.Label lbldeter 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Caption         =   "IMPUREZAS DIVERSAS"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   300
      Index           =   3
      Left            =   6300
      TabIndex        =   107
      Top             =   420
      Width           =   6165
   End
   Begin VB.Label lbldeter 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Caption         =   "IMPUREZAS CONSTITUIDAS POR GRANOS"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   300
      Index           =   0
      Left            =   60
      TabIndex        =   48
      Top             =   840
      Width           =   6165
   End
   Begin VB.Label lbltitulo 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      Caption         =   "Determinaciones Organoleptico"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   330
      Left            =   60
      TabIndex        =   47
      Top             =   30
      Width           =   12405
   End
End
Attribute VB_Name = "frmOrganoleptico"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim granu As Long
    
'Private WithEvents TecladoNumerico As frmTecladoNumerico
'Private blnTecladoNumericoPrimeraVez As Boolean
Private mvarintIndiceValor As Integer
'Private blnEsTablet As Boolean





Private Sub cmdObservador_Click()
Dim objfrm As New frmObservadorEnsayo
    
    objfrm.ES_CONTROL_EFICACIA = False
    objfrm.MUESTRA_ID = gmuestra ' Id de la muestra
    objfrm.TIPO_DETERMINACION_ENSAYO_ID = CLng(ReadINI(App.Path + "\config.ini", "Alveograma", "Organoleptico")) ' tipo de la Determinacion
    objfrm.DETERMINACION_ENSAYO_ID = gdeterminacion
    objfrm.MUESTRA_CERRADA = (Not cmdok.Enabled)
    objfrm.TIPO_OBSERVACION_ID = MC_TIPOS_OBSERVACION.MCTO_DETERMINACION
        
    objfrm.Show vbModal
     
    Set objfrm = Nothing
    
End Sub
'Private Sub ConfigurarTablet()
'Set TecladoNumerico = New frmTecladoNumerico
'
'    TecladoNumerico.OcultarConformidad = True
'    TecladoNumerico.posX = Screen.Width - TecladoNumerico.Width - 60
'    TecladoNumerico.posY = Me.top
'
'    blnEsTablet = pc_es_tablet
'
'    If blnEsTablet Then
'
'        blnTecladoNumericoPrimeraVez = True
'
'        val(0).Locked = True
'        val(1).Locked = True
'        val(2).Locked = True
'        val(3).Locked = True
'        val(4).Locked = True
'        val(5).Locked = True
'        val(6).Locked = True
'        val(7).Locked = True
'        val(8).Locked = True
'        val(9).Locked = True
'        val(10).Locked = True
'        val(11).Locked = True
'        val(12).Locked = True
'        val(13).Locked = True
'        val(14).Locked = True
'        val(15).Locked = True
'
'
'        Me.Left = 0
'    End If
'End Sub
Private Function getCabecera() As String
Select Case mvarintIndiceValor
    Case 0
        getCabecera = Label1(25)
    Case 1, 2, 3, 4, 5
        getCabecera = lbldeter(0).Caption
    Case 6, 7, 8, 9, 10, 11
        getCabecera = lbldeter(3).Caption
    Case 12, 13, 17, 18, 19
        getCabecera = lbldeter(2).Caption
    Case 14, 15, 16
        getCabecera = lbldeter(4).Caption
    Case Else
        getCabecera = ""
    End Select


End Function
Private Function getSubCabecera() As String


getSubCabecera = Label1(100 + mvarintIndiceValor)

End Function
'Private Sub MostrarTecladoNumerico()
'
' If Not blnEsTablet Then Exit Sub
'
''MsgBox blnTecladoNumericoPrimeraVez
'    If blnTecladoNumericoPrimeraVez Then
'        blnTecladoNumericoPrimeraVez = False
'        TecladoNumerico.TextoInicial = val(mvarintIndiceValor).Text
'        TecladoNumerico.cabecera = getCabecera()
'        TecladoNumerico.Subcabecera = getSubCabecera()
'        If Not TecladoNumerico.Visible Then
'            TecladoNumerico.Show 1
'        End If
'    End If
'
'End Sub
Private Sub cmdCalcular_Click()
    On Error Resume Next
    ' Peso de la muestra
    If val(0) = "" Or IsNumeric(val(0)) = False Then
        MsgBox "Debe indicar el Peso de Muestra.", vbInformation, App.Title
        val(0).SetFocus
        Exit Sub
    End If
    ' Peso de los 1000 gramos
    If val(19).Text <> "" Then
        If IsNumeric(val(19)) = False Then
            MsgBox "Ha introducido un valor no númerico en el Peso de los 1000 granos.", vbInformation, App.Title
            val(19).SetFocus
            Exit Sub
        End If
    End If
    calcTotalGranos
    calcTotalDiversas
    calcTotalMF
    ' Validar campos numericos
    Dim i As Integer
    For i = 1 To 19
        If Trim(val(i).Text) <> "" Then
            If IsNumeric(val(i)) = False Then
                MsgBox "Ha introducido un valor no númerico.", vbInformation, App.Title
                val(i).SetFocus
                Exit Sub
            Else
                por(i) = Format((CSng(val(i)) / CSng(val(0)) * 100), "#0.00")
            End If
        End If
    Next
    ' ATACADOS
    If Trim(total(0)) = "" Then
        total(0) = formatear("0", 5, 4)
    End If
    If Trim(total(2)) = "" Then
        total(2) = formatear("0", 5, 4)
    End If
    If Trim(total(4)) = "" Then
        total(4) = formatear("0", 5, 4)
    End If
    If Trim(val(17)) = "" Then
        val(17) = formatear("0", 5, 4)
    End If
    If Trim(val(18)) = "" Then
        val(18) = formatear("0", 5, 4)
    End If
    total(1) = formatear((CSng(total(0)) / CSng(val(0)) * 100), 2, 2)
    total(3) = formatear((CSng(total(2)) / CSng(val(0)) * 100), 2, 2)
    total(5) = formatear((CSng(total(4)) / CSng(val(0)) * 100), 2, 2)
    por(17) = formatear((CSng(val(17)) / CSng(val(0)) * 100), 2, 2)
    por(18) = formatear((CSng(val(18)) / CSng(val(0)) * 100), 2, 2)
  
    por(19) = formatear(CSng(total(1)) + CSng(total(3)) + CSng(total(5)) + CSng(por(17)) + CSng(por(18)), 3, 2)
    ' Total Cortagrano
    Dim total_cortagrano As Single
    Dim total_cortagrano_p As Single
    
    Dim datos As Boolean
    datos = False
    For i = 14 To 16
      If Trim(val(i)) <> "" Then
        total_cortagrano = total_cortagrano + CSng(val(i))
        datos = True
      End If
    Next
    If datos = True And total_cortagrano <> 0 Then
        For i = 14 To 16
          If Trim(val(i)) <> "" Then
            If total_cortagrano <> 0 Then
                por(i) = formatear(str(CSng(val(i)) * 100 / total_cortagrano), 3, 2)
                total_cortagrano_p = total_cortagrano_p + por(i)
            End If
          End If
        Next
        If total_cortagrano <> 50 Then
            MsgBox "Los datos introducidos del Examen al Cortagrano no suman 50 granos.", vbInformation, App.Title
            Exit Sub
        End If
        If total_cortagrano_p <> 100 Then
            MsgBox "Los datos introducidos del Examen al Cortagrano no son correctos.", vbInformation, App.Title
            Exit Sub
        End If
    End If
    ' Total
    If CSng(por(19)) > 100 Then
        MsgBox "El resultado supera el 100%", vbCritical, App.Title
        por(19) = ""
    End If
End Sub

Private Sub cmdcancel_Click()
    Unload Me
End Sub

Private Sub cmdok_Click()
    On Error GoTo fallo
    cmdCalcular_Click
    Dim Msg As String
    If granu = 0 Then
        Msg = "Va a introdicir los datos de la granulometria. ¿Esta seguro?"
    Else
        Msg = "Va a modificar los datos de la granulometria. ¿Esta seguro?"
    End If
    If MsgBox(Msg, vbYesNo + vbQuestion, App.Title) = vbYes Then
        Dim ogran As New clsGranulometrias
        With ogran
         .setDETERMINACION_ID = gdeterminacion
         .setMUESTRA_ID = gmuestra
         If Trim(val(19)) <> "" Then
             .setPESO_100 = Replace(val(19), ",", ".")
         Else
            .setPESO_100 = 0
         End If
         If Trim(val(17)) <> "" Then
            .setPARTIDOS = Replace(val(17), ",", ".")
         Else
            .setPARTIDOS = 0
         End If
         If Trim(val(1)) <> "" Then
            .setMERMADOS = Replace(val(1), ",", ".")
        Else
            .setMERMADOS = 0
        End If
         If Trim(val(12)) <> "" Then
             .setOTROS_CEREALES = Replace(val(2), ",", ".")
        Else
            .setOTROS_CEREALES = 0
        End If
        If Trim(val(3)) <> "" Then
         .setTRIGO_BLANDO = Replace(val(3), ",", ".")
        Else
            .setTRIGO_BLANDO = 0
        End If
        If Trim(val(4)) <> "" Then
             .setATACADOS = Replace(val(4), ",", ".")
        Else
            .setATACADOS = 0
        End If
        If Trim(val(5)) <> "" Then
             .setCON_COLORACION = Replace(val(5), ",", ".")
        Else
            .setCON_COLORACION = 0
        End If
         If Trim(val(12)) <> "" Then
         .setMACULADOS = Replace(val(12), ",", ".")
            Else
            .setMACULADOS = 0
        End If
         If Trim(val(13)) <> "" Then
         .setFUSARIADOS = Replace(val(13), ",", ".")
        Else
        .setFUSARIADOS = 0
    End If
         If Trim(val(18)) <> "" Then
         .setGERMINADOS = Replace(val(18), ",", ".")
        Else
            .setGERMINADOS = 0
        End If
         If Trim(val(6)) <> "" Then
         .setSEM_NOCIVAS = Replace(val(6), ",", ".")
            Else
            .setSEM_NOCIVAS = 0
        End If
         If Trim(val(7)) <> "" Then
         .setSEM_NO_NOCIVAS = Replace(val(7), ",", ".")
         Else
         .setSEM_NO_NOCIVAS = 0
        End If
                 If Trim(val(8)) <> "" Then
         .setAVERIADOS = Replace(val(8), ",", ".")
        Else
            .setAVERIADOS = 0
        End If
        If Trim(val(9)) <> "" Then
         .setOTROS_DANADOS = Replace(val(9), ",", ".")
        Else
        .setOTROS_DANADOS = 0
        End If
         If Trim(val(10)) <> "" Then
         .setPMD = Replace(val(10), ",", ".")
         Else
         .setPMD = 0
         End If
         If Trim(val(11)) <> "" Then
         .setINSECTOS = Replace(val(11), ",", ".")
         Else
         .setINSECTOS = 0
         End If
         If Trim(val(14)) <> "" Then
         .setVITREOS = Replace(val(14), ",", ".")
         Else
         .setVITREOS = 0
         End If
         If Trim(val(15)) <> "" Then
         .setBERRENDOS = Replace(val(15), ",", ".")
         Else
         .setBERRENDOS = 0
         End If
         If Trim(val(16)) <> "" Then
         .setHARINOSOS = Replace(val(16), ",", ".")
         Else
         .setHARINOSOS = 0
         End If
         If Trim(val(0)) <> "" Then
         .setPESO_MUESTRA = Replace(val(0), ",", ".")
         Else
         .setPESO_MUESTRA = 0
         End If
         .setOBSERVACIONES_DET = val(20)
         .setES_DUPLICADO = 0
         If granu = 0 Then
             .InsertarGranulometria
         Else
            .Modificar (granu)
         End If
        End With
        Dim oDeter As New clsDeterminaciones
        Dim odd As New clsDatos_determinaciones
        Dim i As Integer
        ' Almacenar Datos Determinaciones
        If odd.CARGAR(gdeterminacion, 331) = True Then
            odd.setVALOR_1 = Replace(por(19), ",", ".")
            odd.Insertar_Valores
        End If
        ' Almacena determinacion (Solucion)
        oDeter.setRESULTADO = Replace(por(19), ",", ".")
        oDeter.setDIF_DUPLICADOS = ""
        oDeter.setFECHA = Format(Date, "yyyy-mm-dd")
        oDeter.setHORA = Format(Time, "hh:mm")
        oDeter.setEMPLEADO_ID = usuario.getID_EMPLEADO
        oDeter.InsertarSolucion (gdeterminacion)
        Unload Me
    End If
    Exit Sub
fallo:
    MsgBox "Error al insertar los datos (granulometria)", vbCritical, Err.Description
End Sub

Private Sub Form_Activate()

    mvarintIndiceValor = 0
    
'    MostrarTecladoNumerico
    
End Sub

Private Sub Form_Load()
    log (Me.Name)
    cargar_botones Me
    Dim i As Integer
    For i = 1 To 19
        por(i).Enabled = False
    Next
    For i = 0 To 5
        total(i).Enabled = False
    Next
    ' Título
    Dim oMuestra As New clsMuestra
    lbltitulo = "Análisis ORGANOLEPTICO de la muestra : " & Trim(str(gmuestra)) & " (" & oMuestra.CodigoParticular(CLng(gmuestra)) & ")"
    Me.Caption = lbltitulo
    Set oMuestra = Nothing
    ' Comprobar si ya existe
    Dim ogran As New clsGranulometrias
    granu = ogran.ComprobarGranulometria(gmuestra, gdeterminacion)
    If granu <> 0 Then
        With ogran
         .CargarGranulometria (granu)
         val(0) = .getPESO_MUESTRA
         val(1) = .getMERMADOS
         val(2) = .getOTROS_CEREALES
         val(3) = .getTRIGO_BLANDO
         val(4) = .getATACADOS
         val(5) = .getCON_COLORACION
         val(6) = .getSEM_NOCIVAS
         val(7) = .getSEM_NO_NOCIVAS
         val(8) = .getAVERIADOS
         val(9) = .getOTROS_DANADOS
         val(10) = .getPMD
         val(11) = .getINSECTOS
         val(12) = .getMACULADOS
         val(13) = .getFUSARIADOS
         val(14) = .getVITREOS
         val(15) = .getBERRENDOS
         val(16) = .getHARINOSOS
         val(17) = .getPARTIDOS
         val(18) = .getGERMINADOS
         val(19) = .getPESO_100
         val(20) = .getOBSERVACIONES_DET
         For i = 0 To 20
           If i < 14 Or i > 16 Then
            If val(i).Text <> "" Then
                val(i) = formatear(val(i), 5, 2)
            End If
           End If
         Next
         cmdCalcular_Click
'         cmdOk.Visible = False
        End With
    End If
    Set ogran = Nothing
    
'    ConfigurarTablet
    If oMuestra.CargaMuestra(gmuestra) Then
        proteger_campos oMuestra.getCERRADA
    End If
    
End Sub


'Private Sub TecladoNumerico_Change(ByVal res As String)
'    Dim iCont As Integer
'    For iCont = 0 To 19
'        val(iCont).BackColor = vbWhite
'    Next iCont
'
'    val(mvarintIndiceValor).BackColor = &H80C0FF
'    val(mvarintIndiceValor).Text = res
'End Sub
'
'Private Sub TecladoNumerico_Salir()
'    blnTecladoNumericoPrimeraVez = False
'    'cmdCalcular_Click
'    cmdCalcular.SetFocus
'End Sub
'
'
'Private Sub TecladoNumerico_SiguienteElemento(cabecera As String, Subcabecera As String, RESULTADO As String, fecha As String, CONFORME As Integer, Cerrar As Boolean, desestimarEvento As Boolean)
'If mvarintIndiceValor = 16 Then
'    Cerrar = True
'    cmdCalcular_Click
'Else
'    mvarintIndiceValor = mvarintIndiceValor + 1
'    cabecera = getCabecera
'    Subcabecera = getSubCabecera
'    RESULTADO = val(mvarintIndiceValor).Text
'End If
'
'End Sub

Private Sub val_GotFocus(Index As Integer)
    val(Index).BackColor = &H80C0FF
    val(Index).SelStart = 0
    val(Index).SelLength = Len(val(Index))
    
    mvarintIndiceValor = Index
    
'    blnTecladoNumericoPrimeraVez = True
'    MostrarTecladoNumerico
    
End Sub

Private Sub val_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
     Case 40
       If Index = 20 Then
        val(0).SetFocus
       Else
        SendKeys "{Tab}", True
       End If
       KeyAscii = 0 ' Para evitar el "bip" del sistema
     Case 38
       If Index = 0 Then
        val(20).SetFocus
       Else
        SendKeys "+{Tab}", True
       End If
       KeyAscii = 0 ' Para evitar el "bip" del sistema
     Case 27
        cmdcancel_Click
    End Select
End Sub

Private Sub val_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = 13 Then
       If Index = 20 Then
        val(0).SetFocus
        KeyAscii = 0 ' Para evitar el "bip" del sistema
       Else
        SendKeys "{Tab}", True
        KeyAscii = 0 ' Para evitar el "bip" del sistema
       End If
    End If
    ' Escribir ',' al pulsar '.'
    If KeyAscii = 46 Then
         KeyAscii = 44
    End If
End Sub

Private Sub val_LostFocus(Index As Integer)
    val(Index).BackColor = vbWhite
    If Index < 14 Or Index > 16 Then
        If val(Index).Text <> "" Then
            val(Index) = formatear(val(Index), 5, 2)
        End If
    End If
End Sub

Public Sub calcTotalGranos()
    Dim i As Integer
    Dim tot As Single
    For i = 1 To 5
      If Trim(val(i)) <> "" Then
        tot = tot + CSng(val(i))
      End If
    Next
    total(0).Text = Format(tot, "####0.0000")
End Sub

Public Sub calcTotalDiversas()
    Dim i As Integer
    Dim tot As Single
    For i = 6 To 11
      If Trim(val(i)) <> "" Then
        tot = tot + CSng(val(i))
      End If
    Next
    total(2).Text = Format(tot, "####0.0000")
End Sub

Public Sub calcTotalMF()
    Dim i As Integer
    Dim tot As Single
    For i = 12 To 13
      If Trim(val(i)) <> "" Then
        tot = tot + CSng(val(i))
      End If
    Next
    total(4).Text = Format(tot, "####0.0000")
End Sub

Private Sub proteger_campos(CERRADA As Integer)
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
    If CERRADA = 1 Then
        cmdok.Enabled = False
        cmdCalcular.Enabled = False
    Else
        cmdok.Enabled = True
        cmdCalcular.Enabled = True
    End If
End Sub
