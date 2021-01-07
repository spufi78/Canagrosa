VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmEdicionFamiliaEquipo 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Áreas metrológicas"
   ClientHeight    =   2895
   ClientLeft      =   3225
   ClientTop       =   3060
   ClientWidth     =   7230
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2895
   ScaleWidth      =   7230
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox txtCodigo 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   900
      TabIndex        =   1
      Top             =   1125
      Width           =   1800
   End
   Begin MSComDlg.CommonDialog cmddlg 
      Left            =   3360
      Top             =   2145
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
      DialogTitle     =   "Abrir Imagen Equipo"
   End
   Begin VB.TextBox txtNombre 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   900
      TabIndex        =   0
      Top             =   810
      Width           =   6165
   End
   Begin VB.CommandButton cmdCambiarIcono 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Cambiar Icono"
      Height          =   480
      Left            =   1485
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   1515
      Width           =   1245
   End
   Begin VB.CommandButton cmdSalir 
      BackColor       =   &H00E0E0E0&
      Caption         =   "ESC-Salir"
      Height          =   870
      Left            =   6060
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   1995
      Width           =   1050
   End
   Begin VB.CommandButton cmdok 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Aceptar"
      Height          =   870
      Left            =   4980
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   1995
      Width           =   1050
   End
   Begin VB.Label Label3 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Código"
      Height          =   195
      Left            =   135
      TabIndex        =   9
      Top             =   1185
      Width           =   735
   End
   Begin VB.Label Label2 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Icono"
      Height          =   195
      Left            =   120
      TabIndex        =   8
      Top             =   1665
      Width           =   735
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Nombre"
      Height          =   195
      Left            =   120
      TabIndex        =   7
      Top             =   870
      Width           =   735
   End
   Begin VB.Image imgFamilia 
      Height          =   480
      Left            =   900
      Picture         =   "frmEdicionFamiliaEquipo.frx":0000
      Stretch         =   -1  'True
      Top             =   1515
      Width           =   480
   End
   Begin VB.Label lbltitulo 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "Áreas metrológicas de equipos"
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
      TabIndex        =   6
      Top             =   30
      Width           =   3270
   End
   Begin VB.Label lbltitulo 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "Detalle de Áreas metrológicas"
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   1
      Left            =   60
      TabIndex        =   5
      Top             =   345
      Width           =   2100
   End
   Begin VB.Shape fondo 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   645
      Left            =   0
      Top             =   0
      Width           =   7245
   End
End
Attribute VB_Name = "frmEdicionFamiliaEquipo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private mvarblnResultado As Boolean
Private mvarenumTipoEdicion As enumTipoEdicion
Private mvarobjFamiliaEquipo As clsFamiliasEquipos

Private Sub cmdCambiarIcono_Click()
On Error GoTo cmdCambiarIcono_Click_Error
'    Dim objImg As clsArchivoAdjunto
'    Set objImg = mvarobjFamiliaEquipo.getICONO
    
    With cmddlg
        .InitDir = App.Path & "\"
        .ShowOpen
        
        If Trim(.FileName) <> "" Then
            
            If LCase(Split(.FileName, ".")(UBound(Split(.FileName, ".")))) <> "jpg" Then
                MsgBox "La imagen para el equipo debe ser estar en formato JPG", vbInformation, "Cambiar Imagen Equipo"
                Exit Sub
            End If
            
'            objImg.setRUTA_TEMPORAL = .FileName
            Dim oD As New clsDocumentacion
            oD.EliminarEquipoTipo mvarobjFamiliaEquipo.getID, 3, 0
            If oD.SubirEquipo(mvarobjFamiliaEquipo.getID, 3, 0, 0, .FileName, .FileTitle) = "" Then
                Set imgFamilia.Picture = LoadPicture(.FileName)
            End If
            Set oD = Nothing
        End If
    End With
    
'    Set mvarobjFamiliaEquipo.setICONO = objImg

On Error GoTo 0
    Exit Sub
cmdCambiarIcono_Click_Error:
End Sub

Private Sub Form_Unload(Cancel As Integer)

    Set mvarobjFamiliaEquipo = Nothing

End Sub


Private Sub OpcionesEdicion()
If mvarenumTipoEdicion = visualizar Then
    txtNombre.Locked = True
    txtCodigo.Locked = True
    cmdCambiarIcono.Enabled = False
    cmdok.Enabled = False
End If
End Sub

Public Property Get TipoEdicion() As enumTipoEdicion

   On Error GoTo TipoEdicion_Error

    TipoEdicion = mvarenumTipoEdicion

   On Error GoTo 0
   Exit Property

TipoEdicion_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure TipoEdicion of Formulario frmEdicionFamiliaEquipo"

End Property

Public Property Let TipoEdicion(ByVal enumTipoEdicion As enumTipoEdicion)

   On Error GoTo TipoEdicion_Error

    mvarenumTipoEdicion = enumTipoEdicion

   On Error GoTo 0
   Exit Property

TipoEdicion_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure TipoEdicion of Formulario frmEdicionFamiliaEquipo"

End Property


'Public Property Get AccionInmediata() As clsProcNcAccionInmediata
'
'    Set AccionInmediata = mvarobjAccionInmediata
'
'End Property
'
'Public Property Set AccionInmediata(ByRef valor As clsProcNcAccionInmediata)
'
'    Set mvarobjAccionInmediata = valor
'
'End Property

Public Property Get RESULTADO() As Boolean

    RESULTADO = mvarblnResultado

End Property

Public Property Let RESULTADO(ByVal blnResultado As Boolean)

    mvarblnResultado = blnResultado

End Property

Private Sub cmdok_Click()

If Trim(txtNombre.Text) = "" Then
    MsgBox "Ha de establecer un Nombre para la familia", vbInformation, "Familias de equipos"
    Exit Sub
End If
If Trim(txtCodigo.Text) = "" Then
    MsgBox "Ha de establecer un Codigo para la familia", vbInformation, "Familias de equipos"
    Exit Sub
End If

mvarobjFamiliaEquipo.setNOMBRE = txtNombre.Text
mvarobjFamiliaEquipo.setCODIGO = txtCodigo.Text

If mvarenumTipoEdicion = Alta Then
    Call mvarobjFamiliaEquipo.Insertar
Else
    Call mvarobjFamiliaEquipo.Modificar
End If


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
    
    If mvarenumTipoEdicion = Alta Then
        Set mvarobjFamiliaEquipo = New clsFamiliasEquipos
        mvarobjFamiliaEquipo.setID_AUX = -2
        Exit Sub
    End If
    
    ' PresentaDatos
    txtNombre.Text = mvarobjFamiliaEquipo.getNOMBRE
    txtCodigo = mvarobjFamiliaEquipo.getCODIGO
'    If Trim(mvarobjFamiliaEquipo.getICONO) <> "" Then
        Dim oD As New clsDocumentacion
        Set imgFamilia.Picture = LoadPicture(oD.CargarEquipo(mvarobjFamiliaEquipo.getID, 3, 0, 0, False))
        Set oD = Nothing
'    End If
    
    OpcionesEdicion
    
End Sub




Public Property Get FamiliaEquipo() As clsFamiliasEquipos

    Set FamiliaEquipo = mvarobjFamiliaEquipo

End Property

Public Property Set FamiliaEquipo(objFamiliaEquipo As clsFamiliasEquipos)

    Set mvarobjFamiliaEquipo = objFamiliaEquipo

End Property
