VERSION 5.00
Begin VB.Form frmBarraAcceso 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   Caption         =   "Trabajo Diario"
   ClientHeight    =   5640
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   1575
   ClipControls    =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5640
   ScaleWidth      =   1575
   ShowInTaskbar   =   0   'False
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0FF&
      Caption         =   "P.N.C"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   135
      TabIndex        =   2
      Top             =   5220
      Width           =   1320
   End
   Begin VB.Image Image3 
      Appearance      =   0  'Flat
      Height          =   1485
      Left            =   180
      Picture         =   "frmBarraAcceso.frx":0000
      Stretch         =   -1  'True
      Top             =   3735
      Width           =   1260
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0FF&
      Caption         =   "Muestras"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   135
      TabIndex        =   1
      Top             =   3375
      Width           =   1320
   End
   Begin VB.Image Image2 
      Appearance      =   0  'Flat
      Height          =   1485
      Left            =   180
      Picture         =   "frmBarraAcceso.frx":0F74
      Stretch         =   -1  'True
      Top             =   1890
      Width           =   1260
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0FF&
      Caption         =   "Equipos"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   135
      TabIndex        =   0
      Top             =   1530
      Width           =   1320
   End
   Begin VB.Image imgEquipos 
      Appearance      =   0  'Flat
      Height          =   1485
      Left            =   180
      Picture         =   "frmBarraAcceso.frx":177F
      Stretch         =   -1  'True
      Top             =   45
      Width           =   1260
   End
End
Attribute VB_Name = "frmBarraAcceso"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_Load()
    log (Me.Name)
    Me.Left = Screen.Width - Me.Width - frmMenu.ButtonBar.Width - 400
    Me.Top = 200 ' Screen.Height - Me.Height - frmMenu.SmartMenuXP1.Height - frmMenu.StatusBar1.Height - 1300
End Sub

Private Sub Image1_Click()
End Sub

Private Sub imgEquipos_Click()
    frmEquipoCuadernoAvisos.Show
End Sub
