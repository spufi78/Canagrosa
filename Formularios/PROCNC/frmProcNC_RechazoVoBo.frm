VERSION 5.00
Begin VB.Form frmProcNC_AccInmediatas 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Acción Inmediata"
   ClientHeight    =   3765
   ClientLeft      =   2640
   ClientTop       =   3030
   ClientWidth     =   7230
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3765
   ScaleWidth      =   7230
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox txtAccionInmediata 
      Height          =   2025
      Left            =   0
      MultiLine       =   -1  'True
      TabIndex        =   4
      Top             =   810
      Width           =   7215
   End
   Begin VB.CommandButton cmdSalir 
      BackColor       =   &H00E0E0E0&
      Caption         =   "ESC-Salir"
      Height          =   870
      Left            =   6150
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   2880
      Width           =   1050
   End
   Begin VB.CommandButton cmdok 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Aceptar"
      Height          =   870
      Left            =   5070
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   2880
      Width           =   1050
   End
   Begin VB.Label Label12 
      Caption         =   $"frmProcNC_AccInmediatas.frx":0000
      ForeColor       =   &H00FF0000&
      Height          =   615
      Left            =   0
      TabIndex        =   5
      Top             =   2880
      Width           =   4950
      WordWrap        =   -1  'True
   End
   Begin VB.Label lbltitulo 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "Acción Inmediata"
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
      Index           =   0
      Left            =   60
      TabIndex        =   3
      Top             =   30
      Width           =   1800
   End
   Begin VB.Label lbltitulo 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "Describa los datos de la Acción Inmediata, rellenando los siguientes campos."
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   1
      Left            =   60
      TabIndex        =   2
      Top             =   300
      Width           =   5430
   End
   Begin VB.Image imagen 
      Height          =   480
      Left            =   6660
      Picture         =   "frmProcNC_AccInmediatas.frx":00B0
      Top             =   60
      Width           =   480
   End
   Begin VB.Shape fondo 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   825
      Left            =   0
      Top             =   0
      Width           =   7245
   End
End
Attribute VB_Name = "frmProcNC_AccInmediatas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private mvarstrAccionInmediata As String
Private mvarblnResultado As Boolean
Private mvarlngidAccionInmediata As Long

Public Property Get idAccionInmediata() As Long

    idAccionInmediata = mvarlngidAccionInmediata

End Property

Public Property Let idAccionInmediata(ByVal lngidAccionInmediata As Long)

    mvarlngidAccionInmediata = lngidAccionInmediata

End Property

Public Property Get AccionInmediata() As String

    AccionInmediata = mvarstrAccionInmediata

End Property

Public Property Let AccionInmediata(ByVal strAccionInmediata As String)

    mvarstrAccionInmediata = strAccionInmediata

End Property

Public Property Get Resultado() As Boolean

    Resultado = mvarblnResultado

End Property

Public Property Let Resultado(ByVal blnResultado As Boolean)

    mvarblnResultado = blnResultado

End Property

Private Sub cmdok_Click()

mvarstrAccionInmediata = txtAccionInmediata.Text
mvarblnResultado = True

Me.Hide

End Sub

Private Sub cmdSalir_Click()

mvarblnResultado = False

Me.Hide
End Sub


Private Sub Form_Load()
    log (Me.Name)
    cargar_botones Me
    
    
    txtAccionInmediata.Text = mvarstrAccionInmediata
    
End Sub


