VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmEquipoEvento 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Evento de Equipo"
   ClientHeight    =   6450
   ClientLeft      =   3900
   ClientTop       =   3405
   ClientWidth     =   8595
   Icon            =   "frmEquipoEvento.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6450
   ScaleWidth      =   8595
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdVerCVM 
      BackColor       =   &H00E0E0E0&
      Height          =   870
      Left            =   90
      Picture         =   "frmEquipoEvento.frx":6852
      Style           =   1  'Graphical
      TabIndex        =   15
      ToolTipText     =   "Crear documento de solicitud de análisis"
      Top             =   5520
      Visible         =   0   'False
      Width           =   1395
   End
   Begin VB.CommandButton cmdcancel 
      BackColor       =   &H00E0E0E0&
      Caption         =   "ESC-Salir"
      Height          =   870
      Left            =   7620
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   5520
      Width           =   960
   End
   Begin VB.CommandButton cmdok 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Aceptar"
      Height          =   870
      Left            =   6540
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   5520
      Width           =   1020
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   4665
      Left            =   60
      TabIndex        =   0
      Top             =   780
      Width           =   8505
      Begin MSDataListLib.DataCombo cmbEvento 
         Height          =   315
         Left            =   1770
         TabIndex        =   10
         Top             =   30
         Width           =   6615
         _ExtentX        =   11668
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         Style           =   2
         Text            =   ""
      End
      Begin VB.TextBox txtObservaciones 
         Appearance      =   0  'Flat
         Height          =   2640
         Left            =   30
         MaxLength       =   1024
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   5
         Top             =   1965
         Width           =   8415
      End
      Begin MSComCtl2.DTPicker txtFecha 
         Height          =   315
         Left            =   1770
         TabIndex        =   3
         Top             =   750
         Width           =   1755
         _ExtentX        =   3096
         _ExtentY        =   556
         _Version        =   393216
         CustomFormat    =   "dd/MM/yyy HH:mm"
         Format          =   61014019
         CurrentDate     =   40273
         MinDate         =   2
      End
      Begin MSDataListLib.DataCombo cmbRazon 
         Height          =   315
         Left            =   1770
         TabIndex        =   11
         Top             =   390
         Width           =   6615
         _ExtentX        =   11668
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         Style           =   2
         Text            =   ""
      End
      Begin MSDataListLib.DataCombo cmbUsuario 
         Height          =   315
         Left            =   1770
         TabIndex        =   16
         Top             =   1125
         Width           =   6615
         _ExtentX        =   11668
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         Style           =   2
         Text            =   ""
      End
      Begin VB.Label lblFecha 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   315
         Left            =   1770
         TabIndex        =   12
         Top             =   750
         Visible         =   0   'False
         Width           =   1965
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "Usuario Resp. Evento"
         Height          =   195
         Left            =   60
         TabIndex        =   7
         Top             =   1170
         Width           =   1635
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Observaciones al evento"
         Height          =   195
         Left            =   45
         TabIndex        =   6
         Top             =   1710
         Width           =   2025
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Fecha/Hora"
         Height          =   195
         Left            =   60
         TabIndex        =   4
         Top             =   840
         Width           =   915
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Motivo del Evento"
         Height          =   195
         Left            =   60
         TabIndex        =   2
         Top             =   450
         Width           =   1365
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Evento"
         Height          =   195
         Left            =   60
         TabIndex        =   1
         Top             =   90
         Width           =   735
      End
   End
   Begin VB.Label lbltitulo 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "Historial/Trazabilidad de Eventos de Equipo"
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
      Left            =   90
      TabIndex        =   9
      Top             =   120
      Width           =   4635
   End
   Begin VB.Image imagen 
      Height          =   480
      Left            =   8010
      Picture         =   "frmEquipoEvento.frx":711C
      Top             =   120
      Width           =   480
   End
   Begin VB.Label lblsubtitulo 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "Ventana de gestión Eventos de Equipo"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   90
      TabIndex        =   8
      Top             =   420
      Width           =   2775
   End
   Begin VB.Shape fondo 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   735
      Left            =   0
      Top             =   0
      Width           =   8580
   End
End
Attribute VB_Name = "frmEquipoEvento"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public PK As Long

Private mvarobjEvento As clsEquipoEventos
Private mvarobjEquipo As clsEquipos
Private mvarblnResultado As Boolean

Private Sub PresentarDatos_VerCVM()

    If PK = 0 Then Exit Sub
    
    If mvarobjEvento.getCVM_ID = 0 Then Exit Sub
    
    cmdVerCVM.Visible = True
    
    Select Case mvarobjEvento.getEVENTO_ID
        Case EQUIPOS_EVENTOS.EVT_CALIBRACION_REALIZADA
            cmdVerCVM.Caption = "Ver Calibración"
        Case EQUIPOS_EVENTOS.EVT_VERIFICACION_REALIZADA
            cmdVerCVM.Caption = "Ver Verificación"
        Case EQUIPOS_EVENTOS.EVT_MANTENIMIENTO_REALIZADO
            cmdVerCVM.Caption = "Ver Mantenimiento"
    End Select
    

End Sub

Private Sub cmbEvento_Change()
Call cargar_razones_eventos
End Sub

Private Sub cmdcancel_Click()
    mvarblnResultado = False
    Me.Hide
End Sub

Private Sub cmdVerCVM_Click()

    Select Case mvarobjEvento.getEVENTO_ID
        Case EQUIPOS_EVENTOS.EVT_CALIBRACION_REALIZADA
            ver_calibracion
        Case EQUIPOS_EVENTOS.EVT_VERIFICACION_REALIZADA
            ver_verificacion
        Case EQUIPOS_EVENTOS.EVT_MANTENIMIENTO_REALIZADO
            ver_mantenimiento
    End Select
End Sub

Private Sub Form_Unload(Cancel As Integer)

    Set mvarobjEquipo = Nothing

End Sub


Private Function comprobar_datos() As Boolean
    comprobar_datos = False
    
    Dim strCad As String
    
    strCad = ""
    
    If getDataComboSel(cmbEvento) <= 0 Then
        strCad = vbCrLf & " - Debe indicar el Evento"
    End If
    
    If getDataComboSel(cmbRazon) < 0 Or Not cmbRazon.Enabled Then
        strCad = vbCrLf & " - Debe indicar el Motivo del Evento"
    End If
    
    If Not IsDate(txtFecha.value) Then
        strCad = vbCrLf & " - Debe indicar una fecha válida para el Evento"
    End If
    
    If strCad <> "" Then
        MsgBox "Se han detectado los siguientes errores: " & strCad, vbInformation, "Evento de Equipo"
        comprobar_datos = False
        Exit Function
    End If
    
    comprobar_datos = True
    
End Function


Private Function RecogerDatos()


    With mvarobjEvento
        .setCVM_ID = 0
        .setEQUIPO_ID = mvarobjEquipo.getID_EQUIPO
        .setEVENTO_ID = getDataComboSel(cmbEvento)
        .setRAZON_ID = getDataComboSel(cmbRazon)
        .setOBSERVACIONES = txtObservaciones.Text
        .setUSUARIO_ID = USUARIO.getID_EMPLEADO
    End With
    

End Function


Private Sub cmdok_Click()

    If Not comprobar_datos Then Exit Sub
    
    RecogerDatos
    
    If PK <> 0 Then
                consulta = "UPDATE EQ_EVENTOS SET " & _
                        " EVENTO_ID = " & mvarobjEvento.getEVENTO_ID & "," & _
                        " RAZON_ID = " & mvarobjEvento.getRAZON_ID & "," & _
                        " OBSERVACIONES = '" & txtObservaciones & "'," & _
                        " EQUIPO_ID = " & mvarobjEquipo.getID_EQUIPO & "," & _
                        " USUARIO_ID = " & cmbUsuario.BoundText & "," & _
                        " TS = '" & Format(Left(txtFecha.value, 10), "yyyy-mm-dd") & " " & Right(txtFecha.value, 8) & "'" & _
                " WHERE ID_EVENTOEQUIPO = " & PK
            execute_bd consulta
    Else
        'Genera el evento a partir del equipo para tenerlo todo centralizado
        mvarobjEquipo.generar_evento mvarobjEquipo.getID_EQUIPO, mvarobjEvento.getEVENTO_ID, mvarobjEvento.getRAZON_ID, mvarobjEvento.getOBSERVACIONES
    End If
        
    mvarblnResultado = True
    Me.Hide
End Sub


Private Sub Form_Load()
    log Me.Name
    cargar_botones Me
    
    cargar_combos
    presentar_datos
    OpcionesEdicion
End Sub

Private Sub OpcionesEdicion()
    Dim rs As ADODB.Recordset
    Dim consulta As String
    Dim responsable As Boolean
    responsable = False
    Dim p As String
    p = "idresp=" & USUARIO.getID_EMPLEADO & ";"
    Set rs = datos_bd("SELECT * FROM decodificadora WHERE CODIGO = " & DECODIFICADORA.PROCNC_DEPARTAMENTOS & " AND VALOR = 6 AND PARAMETROS like '%" & p & "%'")
    If rs.RecordCount > 0 Then
        responsable = True
    End If
    'JGM-I
    If USUARIO.getID_EMPLEADO = 48 Then
        responsable = True
    End If
    'JGM-F
    
    If PK <> 0 And Not responsable Then
        cmbUsuario.Enabled = False
        txtFecha.Visible = False
        lblFecha.Visible = True
        cmbEvento.Locked = True
        cmbRazon.Locked = True
        txtObservaciones.Locked = True
        cmdok.Visible = False
    End If
    
End Sub

Private Sub cargar_combos()
    
    Dim oDeco As New clsDecodificadora

    oDeco.cargar_combo cmbEvento, DECODIFICADORA.EQ_EVENTOS
    
    cargar_combo cmbUsuario, New clsUsuarios
    
End Sub


Private Sub presentar_datos()

    Set mvarobjEvento = New clsEquipoEventos
'    Dim oUsuario As New clsUsuarios
    
'    lblUsuario.Caption = usuario.getNOMBRE & " " & usuario.getAPELLIDOS
    cmbUsuario.BoundText = USUARIO.getID_EMPLEADO
    
    txtFecha.value = Now
    lblFecha.Caption = Format(Now, "dd/mm/yyyy Hh:Nn")
    
    If PK = 0 Then
        cmbRazon.Enabled = False
        Exit Sub
    End If


    ' Carga los datos
    mvarobjEvento.Carga PK
    
    ' Presenta los datos
    cmbEvento.BoundText = mvarobjEvento.getEVENTO_ID
    cargar_razones_eventos
    cmbRazon.BoundText = mvarobjEvento.getRAZON_ID
    txtFecha.value = mvarobjEvento.getTS
    lblFecha.Caption = Format(mvarobjEvento.getTS, "dd/mm/yyyy Hh:Nn")
    txtObservaciones.Text = mvarobjEvento.getOBSERVACIONES
    
'    oUsuario.CARGAR mvarobjEvento.getUSUARIO_ID
'    lblUsuario.Caption = oUsuario.getNOMBRE & " " & oUsuario.getAPELLIDOS
    cmbUsuario.BoundText = mvarobjEvento.getUSUARIO_ID
'    Set oUsuario = Nothing
    
    PresentarDatos_VerCVM


End Sub


Private Sub cargar_razones_eventos()

On Error GoTo cargar_razones_eventos_Error

    If CLng(cmbEvento.BoundText) = 0 Then
        cmbRazon.Enabled = False
        Exit Sub
    End If
    
    With cmbRazon
        .Enabled = True
        Set .RowSource = mvarobjEvento.Devolver_Listado_Razones_Evento_Combo(CLng(cmbEvento.BoundText))
        .ListField = "RAZON"
        .DataField = "ID"
        .BoundColumn = "ID"
    End With
    
On Error GoTo 0
    Exit Sub
cargar_razones_eventos_Error:
    cmbRazon.Enabled = False
    
End Sub

Public Property Get EQUIPO() As clsEquipos

    Set EQUIPO = mvarobjEquipo

End Property

Public Property Set EQUIPO(objEquipo As clsEquipos)

    Set mvarobjEquipo = objEquipo

End Property

Public Property Get RESULTADO() As Boolean

    RESULTADO = mvarblnResultado

End Property

Public Property Let RESULTADO(ByVal blnResultado As Boolean)

    mvarblnResultado = blnResultado

End Property

Private Sub ver_calibracion()
Dim objfrm  As New frmEquipoEdicionCalibracion

With objfrm
    Set .EQUIPO = mvarobjEquipo
    .ID = CStr(mvarobjEvento.getCVM_ID)
    .TipoEdicion = visualizar
    
    .Show vbModal
        
End With

Unload objfrm
Set objfrm = Nothing


End Sub

Private Sub ver_verificacion()
Dim objfrm  As New frmEquipoEdicionVerificacion

    With objfrm
        Set .EQUIPO = mvarobjEquipo
        .ID = CStr(mvarobjEvento.getCVM_ID)
        .TipoEdicion = visualizar
        
        .Show vbModal
            
    End With

Unload objfrm
Set objfrm = Nothing

End Sub

Private Sub ver_mantenimiento()
Dim objfrm  As New frmEquipoEdicionMtoFechasEdicion


    objfrm.TipoEdicion = visualizar
    Set objfrm.EQUIPO = mvarobjEquipo
    objfrm.id_mantenimiento = CLng(mvarobjEvento.getCVM_ID)
    
    objfrm.Show vbModal
    
    Unload objfrm
    Set objfrm = Nothing

End Sub
