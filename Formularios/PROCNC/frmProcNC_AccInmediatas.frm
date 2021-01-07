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
      Picture         =   "frmProcNC_AccInmediatas.frx":00AB
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
Private mvarblnResultado As Boolean
Private mvarobjAccionInmediata As clsProcNcAccionInmediata
Private mvarenumTipoEdicion As enumTipoEdicion

Private Sub OpcionesEdicion()
If mvarenumTipoEdicion = visualizar Then
    txtAccionInmediata.Locked = True
    cmdok.Enabled = False
End If
End Sub

Public Property Get TipoEdicion() As enumTipoEdicion

   On Error GoTo TipoEdicion_Error

    TipoEdicion = mvarenumTipoEdicion

   On Error GoTo 0
   Exit Property

TipoEdicion_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure TipoEdicion of Formulario frmProcNC_AccInmediatas"

End Property

Public Property Let TipoEdicion(ByVal enumTipoEdicion As enumTipoEdicion)

   On Error GoTo TipoEdicion_Error

    mvarenumTipoEdicion = enumTipoEdicion

   On Error GoTo 0
   Exit Property

TipoEdicion_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure TipoEdicion of Formulario frmProcNC_AccInmediatas"

End Property


Public Property Get AccionInmediata() As clsProcNcAccionInmediata

    Set AccionInmediata = mvarobjAccionInmediata

End Property

Public Property Set AccionInmediata(ByRef valor As clsProcNcAccionInmediata)

    Set mvarobjAccionInmediata = valor

End Property

Public Property Get Resultado() As Boolean

    Resultado = mvarblnResultado

End Property

Public Property Let Resultado(ByVal blnResultado As Boolean)

    mvarblnResultado = blnResultado

End Property

Private Sub cmdok_Click()

On Error GoTo cmdok_Click_Error
    
If Trim(txtAccionInmediata.Text) = "" Then
    MsgBox "Ha de establecer un texto descriptivo de la acción inmediata", vbInformation, "Accion Inmediata"
    Exit Sub
End If

mvarobjAccionInmediata.setDESCRIPCION = txtAccionInmediata.Text
mvarblnResultado = True

Me.Hide
    
On Error GoTo 0
    G_TRAZABILIDAD_ERROR = G_TRAZABILIDAD_ERROR & vbCrLf & " -> frmProcNC_AccInmediatas.cmdok_Click"
    Exit Sub
cmdok_Click_Error:
    G_TRAZABILIDAD_ERROR = G_TRAZABILIDAD_ERROR & vbCrLf & " -> frmProcNC_AccInmediatas.cmdok_Click"
    error_grave Err.Number & " (" & Err.Description & ") in procedure cmdok_Click of Formulario frmProcNC_AccInmediatas" & vbCrLf & "Trazabilidad del Error: " & G_TRAZABILIDAD_ERROR
    G_TRAZABILIDAD_ERROR = ""

End Sub

Private Sub cmdSalir_Click()

On Error GoTo cmdSalir_Click_Error
    
mvarblnResultado = False

Me.Hide
    
On Error GoTo 0
    G_TRAZABILIDAD_ERROR = G_TRAZABILIDAD_ERROR & vbCrLf & " -> frmProcNC_AccInmediatas.cmdSalir_Click"
    Exit Sub
cmdSalir_Click_Error:
    G_TRAZABILIDAD_ERROR = G_TRAZABILIDAD_ERROR & vbCrLf & " -> frmProcNC_AccInmediatas.cmdSalir_Click"
    error_grave Err.Number & " (" & Err.Description & ") in procedure cmdSalir_Click of Formulario frmProcNC_AccInmediatas" & vbCrLf & "Trazabilidad del Error: " & G_TRAZABILIDAD_ERROR
    G_TRAZABILIDAD_ERROR = ""
End Sub


Private Sub Form_Load()
On Error GoTo Form_Load_Error
    
    log (Me.Name)
    cargar_botones Me
    
    If mvarenumTipoEdicion = ALTA Then
        Set mvarobjAccionInmediata = New clsProcNcAccionInmediata
        mvarobjAccionInmediata.setID_AUX = -2
        Exit Sub
    End If
    
    ' PresentaDatos
    txtAccionInmediata.Text = mvarobjAccionInmediata.getDESCRIPCION
    
    OpcionesEdicion
    
On Error GoTo 0
    G_TRAZABILIDAD_ERROR = G_TRAZABILIDAD_ERROR & vbCrLf & " -> frmProcNC_AccInmediatas.Form_Load"
    Exit Sub
Form_Load_Error:
    G_TRAZABILIDAD_ERROR = G_TRAZABILIDAD_ERROR & vbCrLf & " -> frmProcNC_AccInmediatas.Form_Load"
    error_grave Err.Number & " (" & Err.Description & ") in procedure Form_Load of Formulario frmProcNC_AccInmediatas" & vbCrLf & "Trazabilidad del Error: " & G_TRAZABILIDAD_ERROR
    G_TRAZABILIDAD_ERROR = ""
    
End Sub


