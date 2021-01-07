VERSION 5.00
Begin VB.Form frmCE_Lote_Probeta 
   BackColor       =   &H00C0C0C0&
   Caption         =   "Nuevo Lote de Probetas"
   ClientHeight    =   8325
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9915
   Icon            =   "frmCE_Lote_Probeta.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   8325
   ScaleWidth      =   9915
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdAdjuntos 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Adjuntos"
      Height          =   870
      Left            =   5760
      Style           =   1  'Graphical
      TabIndex        =   37
      ToolTipText     =   "Enviar informe por E-mail"
      Top             =   7380
      Width           =   1875
   End
   Begin VB.Frame Frame2 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   "Datos de las cargas de Rotura"
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
      Height          =   3975
      Left            =   45
      TabIndex        =   50
      Top             =   3330
      Width           =   9825
      Begin VB.CheckBox chkiden 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Caption         =   "79 %"
         ForeColor       =   &H80000008&
         Height          =   240
         Index           =   9
         Left            =   450
         TabIndex        =   19
         Top             =   1755
         Width           =   735
      End
      Begin VB.TextBox txtpor 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Enabled         =   0   'False
         Height          =   330
         Index           =   9
         Left            =   2610
         Locked          =   -1  'True
         TabIndex        =   20
         Top             =   1710
         Width           =   2160
      End
      Begin VB.TextBox txtte 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Enabled         =   0   'False
         Height          =   330
         Index           =   9
         Left            =   6480
         TabIndex        =   21
         Top             =   1710
         Width           =   2160
      End
      Begin VB.CheckBox chkiden 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Caption         =   "81 %"
         ForeColor       =   &H80000008&
         Height          =   240
         Index           =   8
         Left            =   450
         TabIndex        =   13
         Top             =   1035
         Width           =   735
      End
      Begin VB.TextBox txtpor 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Enabled         =   0   'False
         Height          =   330
         Index           =   8
         Left            =   2610
         Locked          =   -1  'True
         TabIndex        =   14
         Top             =   990
         Width           =   2160
      End
      Begin VB.TextBox txtte 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Enabled         =   0   'False
         Height          =   330
         Index           =   8
         Left            =   6480
         TabIndex        =   15
         Top             =   990
         Width           =   2160
      End
      Begin VB.CheckBox chkiden 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Caption         =   "74 %"
         ForeColor       =   &H80000008&
         Height          =   240
         Index           =   7
         Left            =   450
         TabIndex        =   31
         Top             =   3195
         Width           =   735
      End
      Begin VB.TextBox txtpor 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Enabled         =   0   'False
         Height          =   330
         Index           =   7
         Left            =   2610
         Locked          =   -1  'True
         TabIndex        =   32
         Top             =   3150
         Width           =   2160
      End
      Begin VB.TextBox txtte 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Enabled         =   0   'False
         Height          =   330
         Index           =   7
         Left            =   6480
         TabIndex        =   33
         Top             =   3150
         Width           =   2160
      End
      Begin VB.CheckBox chkiden 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Caption         =   "76 %"
         ForeColor       =   &H80000008&
         Height          =   240
         Index           =   6
         Left            =   450
         TabIndex        =   25
         Top             =   2475
         Width           =   735
      End
      Begin VB.TextBox txtpor 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Enabled         =   0   'False
         Height          =   330
         Index           =   6
         Left            =   2610
         Locked          =   -1  'True
         TabIndex        =   26
         Top             =   2430
         Width           =   2160
      End
      Begin VB.TextBox txtte 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Enabled         =   0   'False
         Height          =   330
         Index           =   6
         Left            =   6480
         TabIndex        =   27
         Top             =   2430
         Width           =   2160
      End
      Begin VB.TextBox txtte 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Enabled         =   0   'False
         Height          =   330
         Index           =   5
         Left            =   6480
         TabIndex        =   24
         Top             =   2070
         Width           =   2160
      End
      Begin VB.TextBox txtpor 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Enabled         =   0   'False
         Height          =   330
         Index           =   5
         Left            =   2610
         Locked          =   -1  'True
         TabIndex        =   23
         Top             =   2070
         Width           =   2160
      End
      Begin VB.CheckBox chkiden 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Caption         =   "77 %"
         ForeColor       =   &H80000008&
         Height          =   240
         Index           =   5
         Left            =   450
         TabIndex        =   22
         Top             =   2115
         Width           =   735
      End
      Begin VB.TextBox txtte 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Enabled         =   0   'False
         Height          =   330
         Index           =   4
         Left            =   6480
         TabIndex        =   36
         Top             =   3510
         Width           =   2160
      End
      Begin VB.TextBox txtpor 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Enabled         =   0   'False
         Height          =   330
         Index           =   4
         Left            =   2610
         Locked          =   -1  'True
         TabIndex        =   35
         Top             =   3510
         Width           =   2160
      End
      Begin VB.CheckBox chkiden 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Caption         =   "73 %"
         ForeColor       =   &H80000008&
         Height          =   240
         Index           =   4
         Left            =   450
         TabIndex        =   34
         Top             =   3555
         Width           =   735
      End
      Begin VB.TextBox txtte 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Enabled         =   0   'False
         Height          =   330
         Index           =   3
         Left            =   6480
         TabIndex        =   9
         Top             =   270
         Width           =   2160
      End
      Begin VB.TextBox txtpor 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Enabled         =   0   'False
         Height          =   330
         Index           =   3
         Left            =   2610
         Locked          =   -1  'True
         TabIndex        =   8
         Top             =   270
         Width           =   2160
      End
      Begin VB.CheckBox chkiden 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Caption         =   "90 %"
         ForeColor       =   &H80000008&
         Height          =   240
         Index           =   3
         Left            =   450
         TabIndex        =   7
         Top             =   315
         Width           =   735
      End
      Begin VB.TextBox txtte 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Enabled         =   0   'False
         Height          =   330
         Index           =   2
         Left            =   6480
         TabIndex        =   12
         Top             =   630
         Width           =   2160
      End
      Begin VB.TextBox txtpor 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Enabled         =   0   'False
         Height          =   330
         Index           =   2
         Left            =   2610
         Locked          =   -1  'True
         TabIndex        =   11
         Top             =   630
         Width           =   2160
      End
      Begin VB.CheckBox chkiden 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Caption         =   "85 %"
         ForeColor       =   &H80000008&
         Height          =   240
         Index           =   2
         Left            =   450
         TabIndex        =   10
         Top             =   675
         Width           =   735
      End
      Begin VB.TextBox txtte 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Enabled         =   0   'False
         Height          =   330
         Index           =   1
         Left            =   6480
         TabIndex        =   18
         Top             =   1350
         Width           =   2160
      End
      Begin VB.TextBox txtpor 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Enabled         =   0   'False
         Height          =   330
         Index           =   1
         Left            =   2610
         Locked          =   -1  'True
         TabIndex        =   17
         Top             =   1350
         Width           =   2160
      End
      Begin VB.CheckBox chkiden 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Caption         =   "80 %"
         ForeColor       =   &H80000008&
         Height          =   240
         Index           =   1
         Left            =   450
         TabIndex        =   16
         Top             =   1395
         Width           =   735
      End
      Begin VB.TextBox txtte 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Enabled         =   0   'False
         Height          =   330
         Index           =   0
         Left            =   6480
         TabIndex        =   30
         Top             =   2790
         Width           =   2160
      End
      Begin VB.TextBox txtpor 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Enabled         =   0   'False
         Height          =   330
         Index           =   0
         Left            =   2610
         Locked          =   -1  'True
         TabIndex        =   29
         Top             =   2790
         Width           =   2160
      End
      Begin VB.CheckBox chkiden 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Caption         =   "75 %"
         ForeColor       =   &H80000008&
         Height          =   240
         Index           =   0
         Left            =   450
         TabIndex        =   28
         Top             =   2835
         Width           =   735
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Porcentaje Carga"
         Height          =   195
         Index           =   26
         Left            =   1260
         TabIndex        =   70
         Top             =   1755
         Width           =   1230
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Tiempo de Ensayo"
         Height          =   195
         Index           =   25
         Left            =   4995
         TabIndex        =   69
         Top             =   1755
         Width           =   1320
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Porcentaje Carga"
         Height          =   195
         Index           =   24
         Left            =   1260
         TabIndex        =   68
         Top             =   1035
         Width           =   1230
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Tiempo de Ensayo"
         Height          =   195
         Index           =   23
         Left            =   4995
         TabIndex        =   67
         Top             =   1035
         Width           =   1320
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Porcentaje Carga"
         Height          =   195
         Index           =   22
         Left            =   1260
         TabIndex        =   66
         Top             =   3195
         Width           =   1230
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Tiempo de Ensayo"
         Height          =   195
         Index           =   21
         Left            =   4995
         TabIndex        =   65
         Top             =   3195
         Width           =   1320
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Porcentaje Carga"
         Height          =   195
         Index           =   20
         Left            =   1260
         TabIndex        =   64
         Top             =   2475
         Width           =   1230
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Tiempo de Ensayo"
         Height          =   195
         Index           =   19
         Left            =   4995
         TabIndex        =   63
         Top             =   2475
         Width           =   1320
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Tiempo de Ensayo"
         Height          =   195
         Index           =   18
         Left            =   4995
         TabIndex        =   62
         Top             =   2115
         Width           =   1320
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Porcentaje Carga"
         Height          =   195
         Index           =   17
         Left            =   1260
         TabIndex        =   61
         Top             =   2115
         Width           =   1230
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Tiempo de Ensayo"
         Height          =   195
         Index           =   16
         Left            =   4995
         TabIndex        =   60
         Top             =   3555
         Width           =   1320
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Porcentaje Carga"
         Height          =   195
         Index           =   15
         Left            =   1260
         TabIndex        =   59
         Top             =   3555
         Width           =   1230
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Tiempo de Ensayo"
         Height          =   195
         Index           =   14
         Left            =   4995
         TabIndex        =   58
         Top             =   315
         Width           =   1320
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Porcentaje Carga"
         Height          =   195
         Index           =   13
         Left            =   1260
         TabIndex        =   57
         Top             =   315
         Width           =   1230
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Tiempo de Ensayo"
         Height          =   195
         Index           =   12
         Left            =   4995
         TabIndex        =   56
         Top             =   675
         Width           =   1320
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Porcentaje Carga"
         Height          =   195
         Index           =   11
         Left            =   1260
         TabIndex        =   55
         Top             =   675
         Width           =   1230
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Tiempo de Ensayo"
         Height          =   195
         Index           =   10
         Left            =   4995
         TabIndex        =   54
         Top             =   1395
         Width           =   1320
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Porcentaje Carga"
         Height          =   195
         Index           =   9
         Left            =   1260
         TabIndex        =   53
         Top             =   1395
         Width           =   1230
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Tiempo de Ensayo"
         Height          =   195
         Index           =   8
         Left            =   4995
         TabIndex        =   52
         Top             =   2835
         Width           =   1320
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Porcentaje Carga"
         Height          =   195
         Index           =   7
         Left            =   1260
         TabIndex        =   51
         Top             =   2835
         Width           =   1230
      End
   End
   Begin VB.CommandButton cmdok 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Aceptar"
      Height          =   870
      Left            =   7695
      Style           =   1  'Graphical
      TabIndex        =   38
      Top             =   7380
      Width           =   1050
   End
   Begin VB.CommandButton cmdSalir 
      BackColor       =   &H00E0E0E0&
      Caption         =   "ESC-Salir"
      Height          =   870
      Left            =   8820
      Style           =   1  'Graphical
      TabIndex        =   39
      Top             =   7380
      Width           =   1050
   End
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   "Datos del Lote"
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
      Height          =   2580
      Left            =   45
      TabIndex        =   41
      Top             =   360
      Width           =   9810
      Begin VB.TextBox txtDatos 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   330
         Index           =   6
         Left            =   7515
         TabIndex        =   5
         Top             =   1575
         Visible         =   0   'False
         Width           =   2160
      End
      Begin VB.TextBox txtDatos 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   330
         Index           =   5
         Left            =   1440
         TabIndex        =   6
         Top             =   2025
         Width           =   4095
      End
      Begin VB.TextBox txtDatos 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   330
         Index           =   4
         Left            =   1440
         TabIndex        =   4
         Top             =   1665
         Width           =   2160
      End
      Begin VB.TextBox txtDatos 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   330
         Index           =   3
         Left            =   1440
         TabIndex        =   3
         Top             =   1305
         Width           =   4095
      End
      Begin VB.TextBox txtDatos 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   330
         Index           =   2
         Left            =   1440
         TabIndex        =   2
         Top             =   945
         Width           =   4095
      End
      Begin VB.TextBox txtDatos 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   330
         Index           =   1
         Left            =   1440
         TabIndex        =   1
         Top             =   585
         Width           =   8235
      End
      Begin VB.TextBox txtDatos 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   330
         Index           =   0
         Left            =   1440
         TabIndex        =   0
         Top             =   225
         Width           =   8235
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Espesor"
         Height          =   195
         Index           =   5
         Left            =   6165
         TabIndex        =   49
         Top             =   1620
         Visible         =   0   'False
         Width           =   570
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Material"
         Height          =   195
         Index           =   3
         Left            =   90
         TabIndex        =   48
         Top             =   2070
         Width           =   555
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Nº de Lote"
         Height          =   195
         Index           =   1
         Left            =   90
         TabIndex        =   47
         Top             =   1035
         Width           =   765
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Carga de Rotura"
         Height          =   195
         Index           =   6
         Left            =   90
         TabIndex        =   45
         Top             =   1710
         Width           =   1170
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Nº Informe Rotura"
         Height          =   195
         Index           =   4
         Left            =   90
         TabIndex        =   44
         Top             =   1350
         Width           =   1275
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Iden. Selección"
         Height          =   195
         Index           =   0
         Left            =   90
         TabIndex        =   42
         Top             =   675
         Width           =   1110
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Identificación"
         Height          =   195
         Index           =   2
         Left            =   90
         TabIndex        =   43
         Top             =   300
         Width           =   945
      End
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFC0&
      Caption         =   "Cargas de Rotura"
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
      Height          =   255
      Left            =   45
      TabIndex        =   46
      Top             =   3015
      Width           =   9870
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0FF&
      Caption         =   "Nuevo Lote de Probetas"
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
      Height          =   285
      Index           =   2
      Left            =   45
      TabIndex        =   40
      Top             =   45
      Width           =   9840
   End
End
Attribute VB_Name = "frmCE_Lote_Probeta"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public PK As Long
Private Sub cmdAdjuntos_Click()
'M1126-I
    With frmAdjuntos
        .TOBJETO = TOBJETO.TOBJETO_LOTE_PROBETA
        .COBJETO = PK
        .Show 1
    End With
    Set frmAdjuntos = Nothing
'M1126-F
End Sub

Private Sub chkiden_Click(Index As Integer)
    txtpor(Index).Enabled = chkiden(Index).value
    txtte(Index).Enabled = chkiden(Index).value
End Sub

Private Sub cmdok_Click()
   On Error GoTo cmdok_Click_Error

    If validar = True Then
      Dim LOTE As Long
      Dim oCE_LP As New clsCe_lotes_probetas
      With oCE_LP
        .setIDENTIFICACION = txtDatos(0)
        .setIDENTIFICACION_COMBO = txtDatos(1)
        .setNUMERO_LOTE = txtDatos(2)
        .setNUMERO_INFORME = txtDatos(3)
        .setCARGA_ROTURA = txtDatos(4)
        .setMATERIAL = txtDatos(5)
        .setESPESOR = 0
        .setINDICATIVO_CARGA_75 = chkiden(0).value
        .setINDICATIVO_CARGA_80 = chkiden(1).value
        .setINDICATIVO_CARGA_85 = chkiden(2).value
        .setINDICATIVO_CARGA_90 = chkiden(3).value
        .setINDICATIVO_CARGA_73 = chkiden(4).value
        .setINDICATIVO_CARGA_77 = chkiden(5).value
        
        .setINDICATIVO_CARGA_74 = chkiden(7).value
        .setINDICATIVO_CARGA_76 = chkiden(6).value
        
        .setINDICATIVO_CARGA_81 = chkiden(8).value
        .setINDICATIVO_CARGA_79 = chkiden(9).value
        
        
        If txtte(0) <> "" Then
            .setTE_75 = txtte(0)
        End If
        If txtte(1) <> "" Then
            .setTE_80 = txtte(1)
        End If
        If txtte(2) <> "" Then
            .setTE_85 = txtte(2)
        End If
        If txtte(3) <> "" Then
            .setTE_90 = txtte(3)
        End If
        If txtte(4) <> "" Then
            .setTE_73 = txtte(4)
        End If
        If txtte(5) <> "" Then
            .setTE_77 = txtte(5)
        End If
      
      
        If txtte(7) <> "" Then
            .setTE_74 = txtte(7)
        End If
        If txtte(6) <> "" Then
            .setTE_76 = txtte(6)
        End If
        If txtte(8) <> "" Then
            .setTE_81 = txtte(8)
        End If
        If txtte(9) <> "" Then
            .setTE_79 = txtte(9)
        End If
        
      End With
      If PK = 0 Then
        If MsgBox("Va a introducir un nuevo Lote. ¿Esta seguro?", vbQuestion + vbYesNo, App.Title) = vbYes Then
            LOTE = oCE_LP.Insertar
        Else
            Exit Sub
        End If
      Else
        If MsgBox("Va a modificar el Lote. ¿Esta seguro?", vbQuestion + vbYesNo, App.Title) = vbYes Then
            oCE_LP.Modificar (PK)
        Else
            Exit Sub
        End If
      End If
      If PK = 0 Then
          MsgBox "El Lote se ha introducido correctamente.", vbOKOnly + vbInformation, App.Title
      Else
          MsgBox "El Lote se ha modificado correctamente.", vbOKOnly + vbInformation, App.Title
      End If
      Unload Me
    End If

   On Error GoTo 0
   Exit Sub

cmdok_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure cmdok_Click of Formulario frmCE_Lote_Probeta"
End Sub

Private Sub cmdSalir_Click()
    Unload Me
End Sub
Private Sub Form_Load()
    log (Me.Name)
    cargar_botones Me
    If PK <> 0 Then
        Label1(2) = "Modificación de Lote de Probetas"
        Me.Caption = Label1(2).Caption
        cargar_ficha
    End If
End Sub

Private Sub txtDatos_Change(Index As Integer)
    If Index = 4 Then
        If txtDatos(Index) <> "" Then
            If IsNumeric(txtDatos(Index)) Then
                txtpor(0) = Format(0.75 * txtDatos(Index), "0.00")
                txtpor(1) = Format(0.8 * txtDatos(Index), "0.00")
                txtpor(2) = Format(0.85 * txtDatos(Index), "0.00")
                txtpor(3) = Format(0.9 * txtDatos(Index), "0.00")
                txtpor(4) = Format(0.73 * txtDatos(Index), "0.00")
                txtpor(5) = Format(0.77 * txtDatos(Index), "0.00")
            
                txtpor(6) = Format(0.76 * txtDatos(Index), "0.00")
                txtpor(7) = Format(0.74 * txtDatos(Index), "0.00")
                
                txtpor(8) = Format(0.81 * txtDatos(Index), "0.00")
                txtpor(9) = Format(0.79 * txtDatos(Index), "0.00")
            End If
        End If
    End If
End Sub

Private Sub txtdatos_GotFocus(Index As Integer)
    txtDatos(Index).BackColor = &H80C0FF
End Sub

Private Sub txtDatos_KeyPress(Index As Integer, KeyAscii As Integer)
    If Index = 4 Or Index = 6 Then
        If KeyAscii = 46 Then
             KeyAscii = 44
        End If
    End If
    If KeyAscii = 13 Then
        SendKeys "{Tab}", True
        KeyAscii = 0 ' Para evitar el "bip" del sistema
    End If
End Sub

Private Sub txtdatos_LostFocus(Index As Integer)
    txtDatos(Index).BackColor = vbWhite
End Sub
Public Sub cargar_ficha()
    Dim oCE_LP As New clsCe_lotes_probetas
    With oCE_LP
        If .Carga(PK) = True Then
            txtDatos(0) = .getIDENTIFICACION
            txtDatos(1) = .getIDENTIFICACION_COMBO
            txtDatos(2) = .getNUMERO_LOTE
            txtDatos(3) = .getNUMERO_INFORME
            txtDatos(4) = .getCARGA_ROTURA
            txtDatos(5) = .getMATERIAL
            txtDatos(6) = .getESPESOR
            chkiden(0).value = .getINDICATIVO_CARGA_75
            chkiden(1).value = .getINDICATIVO_CARGA_80
            chkiden(2).value = .getINDICATIVO_CARGA_85
            chkiden(3).value = .getINDICATIVO_CARGA_90
            chkiden(4).value = .getINDICATIVO_CARGA_73
            chkiden(5).value = .getINDICATIVO_CARGA_77
            
            chkiden(7).value = .getINDICATIVO_CARGA_74
            chkiden(6).value = .getINDICATIVO_CARGA_76
            
            chkiden(8).value = .getINDICATIVO_CARGA_81
            chkiden(9).value = .getINDICATIVO_CARGA_79
            txtte(0) = .getTE_75
            txtte(1) = .getTE_80
            txtte(2) = .getTE_85
            txtte(3) = .getTE_90
            txtte(4) = .getTE_73
            txtte(5) = .getTE_77
        
            txtte(7) = .getTE_74
            txtte(6) = .getTE_76
            txtte(8) = .getTE_81
            txtte(9) = .getTE_79
        End If
    End With
End Sub
Public Function validar() As Boolean
    validar = True
    If Trim(txtDatos(0)) = "" Then
        MsgBox "Debe darle una identificación al lote.", vbInformation, App.Title
        validar = False
        Exit Function
    End If
    If Trim(txtDatos(4)) = "" Then
        If Not IsNumeric(txtDatos(4)) Then
            MsgBox "La carga de rotura debe ser numérica.", vbInformation, App.Title
            validar = False
            Exit Function
        End If
    End If
End Function

Private Sub txtte_LostFocus(Index As Integer)
    If txtte(Index) <> "" Then
        If Not IsNumeric(txtte(Index)) Then
            MsgBox "El tiempo de ensayo debe ser numérico.", vbExclamation, App.Title
            txtte(Index).SetFocus
        End If
    End If
End Sub
