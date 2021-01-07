VERSION 5.00
Begin VB.Form frmProcNC_CuestionesIdentProblema 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Indicación Problema Pregunta/Respuesta "
   ClientHeight    =   5355
   ClientLeft      =   3885
   ClientTop       =   1950
   ClientWidth     =   7230
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5355
   ScaleWidth      =   7230
   ShowInTaskbar   =   0   'False
   Begin VB.Frame fraTipoPregunta 
      Caption         =   "Tipo Pregunta"
      Height          =   735
      Left            =   0
      TabIndex        =   13
      Top             =   2250
      Width           =   7215
      Begin VB.OptionButton optTipoRespuesta 
         Caption         =   "Sí/No"
         Height          =   195
         Index           =   2
         Left            =   5820
         TabIndex        =   16
         Top             =   330
         Width           =   975
      End
      Begin VB.OptionButton optTipoRespuesta 
         Caption         =   "Numérico"
         Height          =   195
         Index           =   1
         Left            =   3150
         TabIndex        =   15
         Top             =   330
         Width           =   1455
      End
      Begin VB.OptionButton optTipoRespuesta 
         Caption         =   "Texto"
         Height          =   195
         Index           =   0
         Left            =   390
         TabIndex        =   14
         Top             =   330
         Value           =   -1  'True
         Width           =   1455
      End
   End
   Begin VB.Frame fraPregunta 
      Caption         =   "Pregunta"
      Height          =   1425
      Left            =   0
      TabIndex        =   11
      Top             =   810
      Width           =   7215
      Begin VB.TextBox txtPregunta 
         Height          =   1155
         Left            =   60
         MultiLine       =   -1  'True
         TabIndex        =   12
         Top             =   210
         Width           =   7095
      End
   End
   Begin VB.CommandButton cmdSalir 
      BackColor       =   &H00E0E0E0&
      Caption         =   "ESC-Salir"
      Height          =   870
      Left            =   6150
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   4470
      Width           =   1050
   End
   Begin VB.CommandButton cmdok 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Aceptar"
      Height          =   870
      Left            =   5070
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   4470
      Width           =   1050
   End
   Begin VB.Frame fraRespuesta 
      Caption         =   "Respuesta"
      Height          =   1425
      Index           =   0
      Left            =   0
      TabIndex        =   6
      Top             =   3000
      Width           =   7215
      Begin VB.TextBox txtRespuestaAlfanumerica 
         Height          =   1155
         Left            =   60
         MultiLine       =   -1  'True
         TabIndex        =   7
         Top             =   210
         Width           =   7095
      End
   End
   Begin VB.Frame fraRespuesta 
      Caption         =   "Respuesta"
      Height          =   1425
      Index           =   1
      Left            =   0
      TabIndex        =   4
      Top             =   3000
      Visible         =   0   'False
      Width           =   7215
      Begin VB.TextBox txtRespuestaNumerica 
         Height          =   315
         Left            =   60
         TabIndex        =   5
         Top             =   570
         Width           =   7095
      End
   End
   Begin VB.Frame fraRespuesta 
      Caption         =   "Respuesta"
      Height          =   1425
      Index           =   2
      Left            =   0
      TabIndex        =   8
      Top             =   3000
      Visible         =   0   'False
      Width           =   7215
      Begin VB.OptionButton optRespuestaNo 
         Caption         =   "NO"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1095
         Left            =   3930
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   210
         Width           =   2715
      End
      Begin VB.OptionButton optRespuestaSi 
         Caption         =   "SÍ"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1095
         Left            =   600
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   210
         Value           =   -1  'True
         Width           =   2715
      End
   End
   Begin VB.Label lbltitulo 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "Identificación del Problema"
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
      Width           =   2850
   End
   Begin VB.Label lbltitulo 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "Responda a la pregunta correspondiente"
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   1
      Left            =   60
      TabIndex        =   2
      Top             =   300
      Width           =   2895
   End
   Begin VB.Image imagen 
      Height          =   480
      Left            =   6660
      Picture         =   "frmProcNC_CuestionesIdentProblema.frx":0000
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
Attribute VB_Name = "frmProcNC_CuestionesIdentProblema"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private mvarintTipoEdicion As enumTipoEdicion
Private mvarobjPreguntaRespuesta As clsProcNcPreguntaRespuesta

Private mvarintTipoPR As Integer
Private mvarblnResultado As Boolean


Private Sub Form_Unload(Cancel As Integer)

    Set mvarobjPreguntaRespuesta = Nothing

End Sub


Public Property Get PreguntaRespuesta() As clsProcNcPreguntaRespuesta

    Set PreguntaRespuesta = mvarobjPreguntaRespuesta

End Property

Public Property Set PreguntaRespuesta(objPreguntaRespuesta As clsProcNcPreguntaRespuesta)

    Set mvarobjPreguntaRespuesta = objPreguntaRespuesta

End Property

Private Sub cmdok_Click()
Dim strResp As String
Dim idTipoPR As Integer

On Error GoTo cmdok_Click_Error
    
If Not ComprobarDatos() Then Exit Sub

With mvarobjPreguntaRespuesta
    Call getRESPUESTA(strResp, idTipoPR)
    .setTIPO_PREGUNTA_RESPUESTA = idTipoPR
    
    .setRESPUESTA = strResp
    If .getREQUERIDA = False Then .setPREGUNTA = txtPregunta.Text
End With

mvarblnResultado = True

Me.Hide
    
On Error GoTo 0
    G_TRAZABILIDAD_ERROR = G_TRAZABILIDAD_ERROR & vbCrLf & " -> frmProcNC_CuestionesIdentProblema.cmdok_Click"
    Exit Sub
cmdok_Click_Error:
    G_TRAZABILIDAD_ERROR = G_TRAZABILIDAD_ERROR & vbCrLf & " -> frmProcNC_CuestionesIdentProblema.cmdok_Click"
    error_grave Err.Number & " (" & Err.Description & ") in procedure cmdok_Click of Formulario frmProcNC_CuestionesIdentProblema" & vbCrLf & "Trazabilidad del Error: " & G_TRAZABILIDAD_ERROR
    G_TRAZABILIDAD_ERROR = ""

End Sub

Private Sub cmdSalir_Click()

On Error GoTo cmdSalir_Click_Error
    
mvarblnResultado = False

Me.Hide
    
On Error GoTo 0
    G_TRAZABILIDAD_ERROR = G_TRAZABILIDAD_ERROR & vbCrLf & " -> frmProcNC_CuestionesIdentProblema.cmdSalir_Click"
    Exit Sub
cmdSalir_Click_Error:
    G_TRAZABILIDAD_ERROR = G_TRAZABILIDAD_ERROR & vbCrLf & " -> frmProcNC_CuestionesIdentProblema.cmdSalir_Click"
    error_grave Err.Number & " (" & Err.Description & ") in procedure cmdSalir_Click of Formulario frmProcNC_CuestionesIdentProblema" & vbCrLf & "Trazabilidad del Error: " & G_TRAZABILIDAD_ERROR
    G_TRAZABILIDAD_ERROR = ""
End Sub


Private Sub Form_Load()
On Error GoTo Form_Load_Error
    
    log (Me.Name)
    cargar_botones Me

    If mvarintTipoEdicion = ALTA Then Exit Sub
    
    Call PresentarDatos
    Call OpcionesEdicion
    
On Error GoTo 0
    G_TRAZABILIDAD_ERROR = G_TRAZABILIDAD_ERROR & vbCrLf & " -> frmProcNC_CuestionesIdentProblema.Form_Load"
    Exit Sub
Form_Load_Error:
    G_TRAZABILIDAD_ERROR = G_TRAZABILIDAD_ERROR & vbCrLf & " -> frmProcNC_CuestionesIdentProblema.Form_Load"
    error_grave Err.Number & " (" & Err.Description & ") in procedure Form_Load of Formulario frmProcNC_CuestionesIdentProblema" & vbCrLf & "Trazabilidad del Error: " & G_TRAZABILIDAD_ERROR
    G_TRAZABILIDAD_ERROR = ""
    
End Sub



Public Property Get TipoEdicion() As Integer

    TipoEdicion = mvarintTipoEdicion

End Property

Public Property Let TipoEdicion(ByVal intTipoEdicion As Integer)

    mvarintTipoEdicion = intTipoEdicion

End Property


Private Sub optTipoRespuesta_Click(Index As Integer)
Dim cont As Integer

On Error GoTo optTipoRespuesta_Click_Error
    
mvarintTipoPR = Index

For x = 0 To fraRespuesta.Count - 1
    fraRespuesta(x).Visible = (x = Index)
Next x
    
On Error GoTo 0
    G_TRAZABILIDAD_ERROR = G_TRAZABILIDAD_ERROR & vbCrLf & " -> frmProcNC_CuestionesIdentProblema.optTipoRespuesta_Click"
    Exit Sub
optTipoRespuesta_Click_Error:
    G_TRAZABILIDAD_ERROR = G_TRAZABILIDAD_ERROR & vbCrLf & " -> frmProcNC_CuestionesIdentProblema.optTipoRespuesta_Click"
    error_grave Err.Number & " (" & Err.Description & ") in procedure optTipoRespuesta_Click of Formulario frmProcNC_CuestionesIdentProblema" & vbCrLf & "Trazabilidad del Error: " & G_TRAZABILIDAD_ERROR
    G_TRAZABILIDAD_ERROR = ""
End Sub



Private Sub PresentarDatos()

On Error GoTo PresentarDatos_Error
    
    With mvarobjPreguntaRespuesta
        txtPregunta.Text = .getPREGUNTA
        Call setRESPUESTA(.getRESPUESTA, .getTIPO_PREGUNTA_RESPUESTA)
    End With
    
On Error GoTo 0
    G_TRAZABILIDAD_ERROR = G_TRAZABILIDAD_ERROR & vbCrLf & " -> frmProcNC_CuestionesIdentProblema.PresentarDatos"
    Exit Sub
PresentarDatos_Error:
    G_TRAZABILIDAD_ERROR = G_TRAZABILIDAD_ERROR & vbCrLf & " -> frmProcNC_CuestionesIdentProblema.PresentarDatos"
    error_grave Err.Number & " (" & Err.Description & ") in procedure PresentarDatos of Formulario frmProcNC_CuestionesIdentProblema" & vbCrLf & "Trazabilidad del Error: " & G_TRAZABILIDAD_ERROR
    G_TRAZABILIDAD_ERROR = ""
End Sub

Private Sub setRESPUESTA(ByVal strRespuesta As String, ByVal intTipoRespuesta As Integer)

On Error GoTo setRESPUESTA_Error
    
    optTipoRespuesta(intTipoRespuesta).value = True
    Call optTipoRespuesta_Click(intTipoRespuesta)

    Select Case intTipoRespuesta
        Case 0
            txtRespuestaAlfanumerica.Text = strRespuesta
        Case 1
            txtRespuestaNumerica.Text = strRespuesta
        Case Else
            If Trim(strRespuesta) <> "" Then
                If CInt(strRespuesta) = 1 Then
                    optRespuestaSi.value = True
                Else
                    optRespuestaNo.value = True
                End If
            End If
    End Select
    
On Error GoTo 0
    G_TRAZABILIDAD_ERROR = G_TRAZABILIDAD_ERROR & vbCrLf & " -> frmProcNC_CuestionesIdentProblema.setRESPUESTA"
    Exit Sub
setRESPUESTA_Error:
    G_TRAZABILIDAD_ERROR = G_TRAZABILIDAD_ERROR & vbCrLf & " -> frmProcNC_CuestionesIdentProblema.setRESPUESTA"
    error_grave Err.Number & " (" & Err.Description & ") in procedure setRESPUESTA of Formulario frmProcNC_CuestionesIdentProblema" & vbCrLf & "Trazabilidad del Error: " & G_TRAZABILIDAD_ERROR
    G_TRAZABILIDAD_ERROR = ""

End Sub

Private Sub getRESPUESTA(ByRef strRespuesta As String, ByRef intTipoRespuesta As Integer)
' ATENCION, los valores los pasa por preferencia

On Error GoTo getRESPUESTA_Error
    
    intTipoRespuesta = mvarintTipoPR
    
    Select Case mvarintTipoPR
        Case 0
            strRespuesta = txtRespuestaAlfanumerica.Text
        Case 1
            strRespuesta = txtRespuestaNumerica.Text
        Case Else
            If optRespuestaSi.value = True Then
                strRespuesta = "1"
            Else
                strRespuesta = "0"
            End If
    End Select
    
On Error GoTo 0
    G_TRAZABILIDAD_ERROR = G_TRAZABILIDAD_ERROR & vbCrLf & " -> frmProcNC_CuestionesIdentProblema.getRESPUESTA"
    Exit Sub
getRESPUESTA_Error:
    G_TRAZABILIDAD_ERROR = G_TRAZABILIDAD_ERROR & vbCrLf & " -> frmProcNC_CuestionesIdentProblema.getRESPUESTA"
    error_grave Err.Number & " (" & Err.Description & ") in procedure getRESPUESTA of Formulario frmProcNC_CuestionesIdentProblema" & vbCrLf & "Trazabilidad del Error: " & G_TRAZABILIDAD_ERROR
    G_TRAZABILIDAD_ERROR = ""

End Sub


Private Sub OpcionesEdicion()

On Error GoTo OpcionesEdicion_Error
    
    If mvarobjPreguntaRespuesta.getREQUERIDA Then
        fraPregunta.Enabled = False
        fraTipoPregunta.Enabled = False
        'cmdok.Visible = False
    End If
    
    If mvarintTipoEdicion = visualizar Then
        fraPregunta.Enabled = False
        fraTipoPregunta.Enabled = False
        fraRespuesta(0).Enabled = False
        fraRespuesta(1).Enabled = False
        fraRespuesta(2).Enabled = False
        cmdok.Visible = False
        
    End If
    
    
On Error GoTo 0
    G_TRAZABILIDAD_ERROR = G_TRAZABILIDAD_ERROR & vbCrLf & " -> frmProcNC_CuestionesIdentProblema.OpcionesEdicion"
    Exit Sub
OpcionesEdicion_Error:
    G_TRAZABILIDAD_ERROR = G_TRAZABILIDAD_ERROR & vbCrLf & " -> frmProcNC_CuestionesIdentProblema.OpcionesEdicion"
    error_grave Err.Number & " (" & Err.Description & ") in procedure OpcionesEdicion of Formulario frmProcNC_CuestionesIdentProblema" & vbCrLf & "Trazabilidad del Error: " & G_TRAZABILIDAD_ERROR
    G_TRAZABILIDAD_ERROR = ""
    
End Sub

Private Function ComprobarDatos() As Boolean

    Dim blnRes As Boolean, strCad As String
On Error GoTo ComprobarDatos_Error
    
    blnRes = False
    Dim strres As String, intres As Integer
    
    strCad = ""
    
    If Not mvarobjPreguntaRespuesta.getREQUERIDA Then
        If Trim(txtPregunta.Text) = "" Then
            strCad = strCad & vbCrLf & " - Debe indicar una Pregunta Válida"
            blnRes = True
        End If
    End If
    
    Call getRESPUESTA(strres, intres)
    
    
    
    If Trim(strres) = "" Then
        blnRes = True
        strCad = strCad & vbCrLf & " - Debe indicar una Respuesta Válida"
    End If

    If blnRes Then
        MsgBox "Se han encontrado los siguientes Errores: " & strCad, vbInformation, "Pregunta Identificación Problema"
    End If


    ComprobarDatos = Not blnRes
    
On Error GoTo 0
    G_TRAZABILIDAD_ERROR = G_TRAZABILIDAD_ERROR & vbCrLf & " -> frmProcNC_CuestionesIdentProblema.ComprobarDatos"
    Exit Function
ComprobarDatos_Error:
    G_TRAZABILIDAD_ERROR = G_TRAZABILIDAD_ERROR & vbCrLf & " -> frmProcNC_CuestionesIdentProblema.ComprobarDatos"
    error_grave Err.Number & " (" & Err.Description & ") in procedure ComprobarDatos of Formulario frmProcNC_CuestionesIdentProblema" & vbCrLf & "Trazabilidad del Error: " & G_TRAZABILIDAD_ERROR
    G_TRAZABILIDAD_ERROR = ""

End Function

Public Property Get Resultado() As Boolean

    Resultado = mvarblnResultado

End Property

Public Property Let Resultado(ByVal blnResultado As Boolean)

    mvarblnResultado = blnResultado

End Property

Private Sub txtRespuestaNumerica_KeyPress(KeyAscii As Integer)
On Error GoTo txtRespuestaNumerica_KeyPress_Error
    
    KeyAscii = KeyAscii_SoloDecimal(txtRespuestaNumerica, KeyAscii)
    
On Error GoTo 0
    G_TRAZABILIDAD_ERROR = G_TRAZABILIDAD_ERROR & vbCrLf & " -> frmProcNC_CuestionesIdentProblema.txtRespuestaNumerica_KeyPress"
    Exit Sub
txtRespuestaNumerica_KeyPress_Error:
    G_TRAZABILIDAD_ERROR = G_TRAZABILIDAD_ERROR & vbCrLf & " -> frmProcNC_CuestionesIdentProblema.txtRespuestaNumerica_KeyPress"
    error_grave Err.Number & " (" & Err.Description & ") in procedure txtRespuestaNumerica_KeyPress of Formulario frmProcNC_CuestionesIdentProblema" & vbCrLf & "Trazabilidad del Error: " & G_TRAZABILIDAD_ERROR
    G_TRAZABILIDAD_ERROR = ""
End Sub


