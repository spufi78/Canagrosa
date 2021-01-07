VERSION 5.00
Begin VB.Form frmProcNC_CambioEstadoRazones 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Rechazo Vº Bº"
   ClientHeight    =   3765
   ClientLeft      =   2640
   ClientTop       =   3030
   ClientWidth     =   7230
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3765
   ScaleWidth      =   7230
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtMotivoRechazo 
      Height          =   1875
      Left            =   0
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   930
      Width           =   7215
   End
   Begin VB.CommandButton cmdSalir 
      BackColor       =   &H00E0E0E0&
      Caption         =   "ESC-Salir"
      Height          =   870
      Left            =   6150
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   2880
      Width           =   1050
   End
   Begin VB.CommandButton cmdok 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Aceptar"
      Height          =   870
      Left            =   5070
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   2880
      Width           =   1050
   End
   Begin VB.Label lbltitulo 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "Rechazar Visto Bueno Calidad"
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
      TabIndex        =   4
      Top             =   30
      Width           =   3180
   End
   Begin VB.Label lbltitulo 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Caption         =   $"frmProcNC_CambioEstadoRazones.frx":0000
      ForeColor       =   &H80000008&
      Height          =   645
      Index           =   1
      Left            =   60
      TabIndex        =   3
      Top             =   300
      Width           =   6345
   End
   Begin VB.Image imagen 
      Height          =   480
      Left            =   6660
      Picture         =   "frmProcNC_CambioEstadoRazones.frx":00C1
      Top             =   60
      Width           =   480
   End
   Begin VB.Shape fondo 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   975
      Left            =   0
      Top             =   0
      Width           =   7245
   End
End
Attribute VB_Name = "frmProcNC_CambioEstadoRazones"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private mvarstrMotivoRechazo As String
Private mvarblnResultado As Boolean
Private mvarlngidMotivoRechazo As Long
Private mvarstrTitulo As String


Public Property Get idMotivoRechazo() As Long

    idMotivoRechazo = mvarlngidMotivoRechazo

End Property

Public Property Let idMotivoRechazo(ByVal lngidMotivoRechazo As Long)

    mvarlngidMotivoRechazo = lngidMotivoRechazo

End Property

Public Property Get MotivoRechazo() As String

    MotivoRechazo = mvarstrMotivoRechazo

End Property

Public Property Let MotivoRechazo(ByVal strMotivoRechazo As String)

    mvarstrMotivoRechazo = strMotivoRechazo

End Property

Public Property Get RESULTADO() As Boolean

    RESULTADO = mvarblnResultado

End Property

Public Property Let RESULTADO(ByVal blnResultado As Boolean)

    mvarblnResultado = blnResultado

End Property

Public Property Let titulo(ByVal dato As String)
    mvarstrTitulo = dato
End Property

Private Sub cmdok_Click()

On Error GoTo cmdok_Click_Error
    
mvarstrMotivoRechazo = txtMotivoRechazo.Text
mvarblnResultado = True

Me.Hide
    
On Error GoTo 0
    G_TRAZABILIDAD_ERROR = G_TRAZABILIDAD_ERROR & vbCrLf & " -> frmProcNC_CambioEstadoRazones.cmdok_Click"
    Exit Sub
cmdok_Click_Error:
    G_TRAZABILIDAD_ERROR = G_TRAZABILIDAD_ERROR & vbCrLf & " -> frmProcNC_CambioEstadoRazones.cmdok_Click"
    error_grave Err.Number & " (" & Err.Description & ") in procedure cmdok_Click of Formulario frmProcNC_CambioEstadoRazones" & vbCrLf & "Trazabilidad del Error: " & G_TRAZABILIDAD_ERROR
    G_TRAZABILIDAD_ERROR = ""

End Sub

Private Sub cmdsalir_Click()

On Error GoTo cmdSalir_Click_Error
    
mvarblnResultado = False

Me.Hide
    
On Error GoTo 0
    G_TRAZABILIDAD_ERROR = G_TRAZABILIDAD_ERROR & vbCrLf & " -> frmProcNC_CambioEstadoRazones.cmdSalir_Click"
    Exit Sub
cmdSalir_Click_Error:
    G_TRAZABILIDAD_ERROR = G_TRAZABILIDAD_ERROR & vbCrLf & " -> frmProcNC_CambioEstadoRazones.cmdSalir_Click"
    error_grave Err.Number & " (" & Err.Description & ") in procedure cmdSalir_Click of Formulario frmProcNC_CambioEstadoRazones" & vbCrLf & "Trazabilidad del Error: " & G_TRAZABILIDAD_ERROR
    G_TRAZABILIDAD_ERROR = ""
End Sub


Private Sub Form_Load()
On Error GoTo Form_Load_Error
    
    log (Me.Name)
    cargar_botones Me
    
    Me.Caption = mvarstrTitulo
    lbltitulo(0).Caption = mvarstrTitulo
    
    txtMotivoRechazo.Text = mvarstrMotivoRechazo
    
On Error GoTo 0
    G_TRAZABILIDAD_ERROR = G_TRAZABILIDAD_ERROR & vbCrLf & " -> frmProcNC_CambioEstadoRazones.Form_Load"
    Exit Sub
Form_Load_Error:
    G_TRAZABILIDAD_ERROR = G_TRAZABILIDAD_ERROR & vbCrLf & " -> frmProcNC_CambioEstadoRazones.Form_Load"
    error_grave Err.Number & " (" & Err.Description & ") in procedure Form_Load of Formulario frmProcNC_CambioEstadoRazones" & vbCrLf & "Trazabilidad del Error: " & G_TRAZABILIDAD_ERROR
    G_TRAZABILIDAD_ERROR = ""
    
End Sub


