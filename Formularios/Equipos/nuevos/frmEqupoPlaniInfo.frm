VERSION 5.00
Begin VB.Form frmEquipoPlanoInfo 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Info Máquina"
   ClientHeight    =   4350
   ClientLeft      =   3135
   ClientTop       =   2250
   ClientWidth     =   8715
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4350
   ScaleWidth      =   8715
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtSerie 
      Height          =   315
      Left            =   4980
      TabIndex        =   4
      Top             =   630
      Width           =   3675
   End
   Begin VB.TextBox txtNombre 
      Height          =   315
      Left            =   4980
      TabIndex        =   3
      Top             =   270
      Width           =   3675
   End
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "Cancelar"
      Height          =   525
      Left            =   7410
      TabIndex        =   1
      Top             =   3750
      Width           =   1245
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "Aceptar"
      Height          =   525
      Left            =   6090
      TabIndex        =   0
      Top             =   3750
      Width           =   1245
   End
   Begin VB.Label lblCap 
      Caption         =   "Nº Serie"
      Height          =   225
      Index           =   1
      Left            =   3270
      TabIndex        =   5
      Top             =   690
      Width           =   1665
   End
   Begin VB.Label lblCap 
      Caption         =   "Nombre de la Maquina"
      Height          =   225
      Index           =   0
      Left            =   3270
      TabIndex        =   2
      Top             =   330
      Width           =   1665
   End
   Begin VB.Image Image1 
      Height          =   2865
      Left            =   90
      Top             =   60
      Width           =   3135
   End
End
Attribute VB_Name = "frmEquipoPlanoInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private mvarlngIndice As Long
Private mvarobjEquipo As clsEquipos

Private Sub Form_Unload(Cancel As Integer)

    Set mvarobjEquipo = Nothing

End Sub


Public Property Get Equipo() As clsEquipos

    Set Equipo = mvarobjEquipo

End Property

Public Property Set Equipo(objEquipo As clsEquipos)

    Set mvarobjEquipo = objEquipo

End Property

Private Sub cmdCancelar_Click()
Me.Hide
End Sub

Private Sub Form_Load()
txtNombre.Text = mvarobjEquipo.getNOMBRE
txtSerie.Text = mvarobjEquipo.getSERIE
End Sub



Public Property Get Indice() As Long

    Indice = mvarlngIndice

End Property

Public Property Let Indice(ByVal lngIndice As Long)

    mvarlngIndice = lngIndice

End Property

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
If UnloadMode = 0 Then Cancel = 1
End Sub

