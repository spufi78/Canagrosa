VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.ocx"
Object = "{CEA3B7F6-2847-4E5E-A551-DB7A62489D44}#39.0#0"; "miCombo.ocx"
Begin VB.Form frmProcNCInvestigacionEscena 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Investigación de la Escena"
   ClientHeight    =   9255
   ClientLeft      =   3495
   ClientTop       =   1230
   ClientWidth     =   10680
   Icon            =   "frmProcNCInvestigacionEscena.frx":0000
   LinkTopic       =   "frmProcNCInvestigacionEscena"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9255
   ScaleWidth      =   10680
   Begin pryCombo.miCombo cmbPersonal 
      Height          =   315
      Left            =   30
      TabIndex        =   1
      Top             =   2250
      Width           =   4635
      _ExtentX        =   8176
      _ExtentY        =   556
   End
   Begin VB.CommandButton cmdcancel 
      BackColor       =   &H00E0E0E0&
      Caption         =   "ESC-Salir"
      Height          =   915
      Left            =   9630
      Style           =   1  'Graphical
      TabIndex        =   25
      Top             =   8310
      Width           =   1020
   End
   Begin VB.TextBox txtRecoleccionDatos 
      Appearance      =   0  'Flat
      Height          =   1515
      Left            =   30
      MaxLength       =   65000
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   330
      Width           =   10605
   End
   Begin VB.CommandButton cmdAdjuntar 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Adjuntos"
      Height          =   945
      Left            =   9780
      Style           =   1  'Graphical
      TabIndex        =   24
      Top             =   7080
      Width           =   855
   End
   Begin VB.CommandButton cmdSubirPrioridad 
      BackColor       =   &H00E0E0E0&
      Height          =   285
      Left            =   4920
      Picture         =   "frmProcNCInvestigacionEscena.frx":08CA
      Style           =   1  'Graphical
      TabIndex        =   21
      ToolTipText     =   "Añadir accesorio"
      Top             =   7020
      Width           =   285
   End
   Begin VB.CommandButton cmdBajarPrioridad 
      BackColor       =   &H00E0E0E0&
      Height          =   285
      Left            =   4920
      Picture         =   "frmProcNCInvestigacionEscena.frx":5F7F
      Style           =   1  'Graphical
      TabIndex        =   22
      ToolTipText     =   "Eliminar accesorio"
      Top             =   7320
      Width           =   285
   End
   Begin VB.TextBox txtSecuencia 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   30
      TabIndex        =   17
      Top             =   6690
      Width           =   4485
   End
   Begin VB.CommandButton cmdAnadirSecuencia 
      BackColor       =   &H00E0E0E0&
      Height          =   285
      Left            =   4590
      Picture         =   "frmProcNCInvestigacionEscena.frx":B63C
      Style           =   1  'Graphical
      TabIndex        =   18
      ToolTipText     =   "Añadir accesorio"
      Top             =   6690
      Width           =   285
   End
   Begin VB.CommandButton cmdEliminarSecuencia 
      BackColor       =   &H00E0E0E0&
      Height          =   285
      Left            =   4920
      Picture         =   "frmProcNCInvestigacionEscena.frx":B861
      Style           =   1  'Graphical
      TabIndex        =   20
      ToolTipText     =   "Eliminar accesorio"
      Top             =   6690
      Width           =   285
   End
   Begin VB.TextBox txtExperiencia 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   2160
      MaxLength       =   65000
      MultiLine       =   -1  'True
      TabIndex        =   16
      Top             =   6030
      Width           =   8415
   End
   Begin VB.TextBox txtCondicionesOperacion 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   2160
      MaxLength       =   65000
      TabIndex        =   11
      Top             =   4380
      Width           =   8415
   End
   Begin VB.TextBox txtCondicionesAmbientales 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   2160
      MaxLength       =   65000
      TabIndex        =   12
      Top             =   4710
      Width           =   8415
   End
   Begin VB.TextBox txtComunicacion 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   2160
      MaxLength       =   65000
      TabIndex        =   13
      Top             =   5040
      Width           =   8415
   End
   Begin VB.TextBox txtCambiosRecientes 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   2160
      MaxLength       =   65000
      TabIndex        =   14
      Top             =   5370
      Width           =   8415
   End
   Begin VB.TextBox txtFormacion 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   2160
      MaxLength       =   65000
      TabIndex        =   15
      Top             =   5700
      Width           =   8415
   End
   Begin VB.CommandButton cmdEliminarEquipo 
      BackColor       =   &H00E0E0E0&
      Enabled         =   0   'False
      Height          =   285
      Left            =   10350
      Picture         =   "frmProcNCInvestigacionEscena.frx":B9F5
      Style           =   1  'Graphical
      TabIndex        =   8
      ToolTipText     =   "Eliminar accesorio"
      Top             =   2280
      Width           =   285
   End
   Begin VB.CommandButton cmdAnadirEquipo 
      BackColor       =   &H00E0E0E0&
      Enabled         =   0   'False
      Height          =   285
      Left            =   10020
      Picture         =   "frmProcNCInvestigacionEscena.frx":BB89
      Style           =   1  'Graphical
      TabIndex        =   6
      ToolTipText     =   "Añadir accesorio"
      Top             =   2280
      Width           =   285
   End
   Begin VB.TextBox txtLocalizacion 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   2160
      MaxLength       =   65000
      TabIndex        =   10
      Top             =   4050
      Width           =   8415
   End
   Begin VB.CommandButton cmdEliminarPersonal 
      BackColor       =   &H00E0E0E0&
      Enabled         =   0   'False
      Height          =   285
      Left            =   5040
      Picture         =   "frmProcNCInvestigacionEscena.frx":BDAE
      Style           =   1  'Graphical
      TabIndex        =   4
      ToolTipText     =   "Eliminar accesorio"
      Top             =   2280
      Width           =   285
   End
   Begin VB.CommandButton cmdAnadirPersonal 
      BackColor       =   &H00E0E0E0&
      Enabled         =   0   'False
      Height          =   285
      Left            =   4710
      Picture         =   "frmProcNCInvestigacionEscena.frx":BF42
      Style           =   1  'Graphical
      TabIndex        =   2
      ToolTipText     =   "Añadir accesorio"
      Top             =   2280
      Width           =   285
   End
   Begin MSComCtl2.DTPicker txtFechaHora 
      Height          =   300
      Left            =   2160
      TabIndex        =   9
      Top             =   3720
      Width           =   1905
      _ExtentX        =   3360
      _ExtentY        =   529
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      CalendarTitleBackColor=   14737632
      CustomFormat    =   "dd/MM/yyyy HH:mm"
      Format          =   53411843
      UpDown          =   -1  'True
      CurrentDate     =   2
      MinDate         =   2
   End
   Begin pryCombo.miCombo cmbEquipos 
      Height          =   315
      Left            =   5400
      TabIndex        =   5
      Top             =   2250
      Width           =   4545
      _ExtentX        =   8017
      _ExtentY        =   556
   End
   Begin MSComctlLib.ListView lstPersonal 
      Height          =   1125
      Left            =   30
      TabIndex        =   3
      Top             =   2580
      Width           =   5295
      _ExtentX        =   9340
      _ExtentY        =   1984
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      AllowReorder    =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   12640511
      BorderStyle     =   1
      Appearance      =   0
      NumItems        =   0
   End
   Begin MSComctlLib.ListView lstEquipos 
      Height          =   1125
      Left            =   5430
      TabIndex        =   7
      Top             =   2580
      Width           =   5205
      _ExtentX        =   9181
      _ExtentY        =   1984
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      AllowReorder    =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   12640511
      BorderStyle     =   1
      Appearance      =   0
      NumItems        =   0
   End
   Begin MSComctlLib.ListView lstSecuencia 
      Height          =   1125
      Left            =   30
      TabIndex        =   19
      Top             =   6990
      Width           =   4845
      _ExtentX        =   8546
      _ExtentY        =   1984
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      AllowReorder    =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   12640511
      BorderStyle     =   1
      Appearance      =   0
      NumItems        =   0
   End
   Begin MSComctlLib.ListView lstDocumentacion 
      Height          =   1125
      Left            =   5250
      TabIndex        =   23
      Top             =   6990
      Width           =   4485
      _ExtentX        =   7911
      _ExtentY        =   1984
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      AllowReorder    =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   12640511
      BorderStyle     =   1
      Appearance      =   0
      NumItems        =   0
   End
   Begin VB.Label lblCap 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      Caption         =   "(Doble Click para abrir archivo)"
      ForeColor       =   &H00FF0000&
      Height          =   195
      Index           =   0
      Left            =   7530
      TabIndex        =   29
      Top             =   6750
      Width           =   2175
   End
   Begin VB.Label lblCampos 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      Caption         =   "Recolección de Datos"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   1
      Left            =   60
      TabIndex        =   30
      Top             =   60
      Width           =   2325
   End
   Begin VB.Label lblCampos 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      Caption         =   "Secuencia"
      Height          =   195
      Index           =   0
      Left            =   90
      TabIndex        =   38
      Top             =   6450
      Width           =   765
   End
   Begin VB.Label lblCampos 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      Caption         =   "Experiencia"
      Height          =   195
      Index           =   24
      Left            =   90
      TabIndex        =   37
      Top             =   6090
      Width           =   825
   End
   Begin VB.Label lblCampos 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      Caption         =   "Condiciones de Operación"
      Height          =   195
      Index           =   22
      Left            =   90
      TabIndex        =   36
      Top             =   4410
      Width           =   1875
   End
   Begin VB.Label lblCampos 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      Caption         =   "Condiciones Ambientales"
      Height          =   195
      Index           =   23
      Left            =   90
      TabIndex        =   35
      Top             =   4740
      Width           =   1770
   End
   Begin VB.Label lblCampos 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      Caption         =   "Comunicación"
      Height          =   195
      Index           =   25
      Left            =   90
      TabIndex        =   33
      Top             =   5070
      Width           =   1005
   End
   Begin VB.Label lblCampos 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      Caption         =   "Cambios Recientes"
      Height          =   195
      Index           =   30
      Left            =   90
      TabIndex        =   32
      Top             =   5400
      Width           =   1365
   End
   Begin VB.Label lblCampos 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      Caption         =   "Formación"
      Height          =   195
      Index           =   31
      Left            =   90
      TabIndex        =   31
      Top             =   5730
      Width           =   735
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Equipos Implicados"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   5430
      TabIndex        =   28
      Top             =   1950
      Width           =   2295
   End
   Begin VB.Label lblCampos 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      Caption         =   "Fecha y Hora"
      Height          =   195
      Index           =   27
      Left            =   90
      TabIndex        =   27
      Top             =   3780
      Width           =   960
   End
   Begin VB.Label lblCampos 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      Caption         =   "Localización"
      Height          =   195
      Index           =   26
      Left            =   90
      TabIndex        =   34
      Top             =   4110
      Width           =   885
   End
   Begin VB.Label Label6 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Personal Implicado"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   60
      TabIndex        =   26
      Top             =   1950
      Width           =   3135
   End
End
Attribute VB_Name = "frmProcNCInvestigacionEscena"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public PK As Long

Private mvarobjProcNC As New clsProcNc
Private RS As ADODB.RecordSet
Private strSql As String

Private mvarblnEditable As Boolean
Private Function guardar_datos() As Boolean
On Error GoTo guardar_datos_Error
    
    mvarobjProcNC.guardar_datos_investigacion_escena txtRecoleccionDatos.Text, _
    txtFechaHora.value, txtLocalizacion.Text, txtCondicionesOperacion.Text, _
    txtCondicionesAmbientales.Text, txtComunicacion.Text, txtCambiosRecientes.Text, _
    txtFormacion.Text, txtExperiencia.Text
    
    guardar_datos = True
    
On Error GoTo 0
    Exit Function
guardar_datos_Error:
    guardar_datos = False
End Function

Private Sub opciones_edicion()

    If Not mvarblnEditable Then
        cmbPersonal.desactivar
        cmbEquipos.desactivar
    End If
        
    txtRecoleccionDatos.Enabled = mvarblnEditable
    'lstPersonal.Enabled = mvarblnEditable
    'lstEquipos.Enabled = mvarblnEditable
    'lstSecuencia.Enabled = mvarblnEditable
    
    cmdAnadirEquipo.Enabled = mvarblnEditable
    cmdAnadirPersonal.Enabled = mvarblnEditable
    cmdEliminarEquipo.Enabled = mvarblnEditable
    cmdEliminarPersonal.Enabled = mvarblnEditable
    cmdAnadirSecuencia.Enabled = mvarblnEditable
    cmdEliminarSecuencia.Enabled = mvarblnEditable
    
    txtSecuencia.Enabled = mvarblnEditable
    txtFechaHora.Enabled = mvarblnEditable
    txtLocalizacion.Enabled = mvarblnEditable
    txtCondicionesAmbientales.Enabled = mvarblnEditable
    txtCondicionesOperacion.Enabled = mvarblnEditable
    txtComunicacion.Enabled = mvarblnEditable
    txtCambiosRecientes.Enabled = mvarblnEditable
    txtFormacion.Enabled = mvarblnEditable
    txtExperiencia.Enabled = mvarblnEditable
    
    cmdAdjuntar.Enabled = mvarblnEditable
End Sub

Private Sub cmdAdjuntar_Click()
    With frmAdjuntos
        .TOBJETO = TOBJETO.TOBJETO_PROCNC_RECOLECCION_DATOS
        .COBJETO = PK
        .Show 1
    End With
    Set frmAdjuntos = Nothing
    Call PresentarDatos_DocumentosAdjuntos
End Sub

Private Sub cmdAnadirEquipo_Click()
    Dim lngid As Long
    lngid = cmbEquipos.getPK_SALIDA
    
    If lngid <= 0 Then Exit Sub
    
    If mvarobjProcNC.anadir_equipamiento_implicado(lngid) Then
        PresentarDatos_EquipamientoImplicado
    End If

End Sub

Private Sub cmdAnadirPersonal_Click()
    
    Dim lngid As Long
    lngid = cmbPersonal.getPK_SALIDA
    If lngid <= 0 Then Exit Sub
    
    If mvarobjProcNC.anadir_personal_implicado(lngid) Then
        PresentarDatos_PersonalImplicado
    End If
End Sub


Private Sub cmdAnadirSecuencia_Click()
    If mvarobjProcNC.anadir_secuencia_investigacion(txtSecuencia.Text) Then
        txtSecuencia.Text = ""
        PresentarDatos_Secuencia
    End If
End Sub

Private Sub cmdBajarPrioridad_Click()
If lstSecuencia.ListItems.Count = 0 Then Exit Sub

Dim lngid As Long

    lngid = lstSecuencia.selectedItem
    'MsgBox lngId
    
    mvarobjProcNC.secuencia_cambiar_orden lngid, lngid + 1
    PresentarDatos_Secuencia

End Sub

Private Sub cmdEliminarEquipo_Click()
If lstEquipos.ListItems.Count = 0 Then Exit Sub

Dim lngid As Long

    lngid = lstEquipos.selectedItem
    'MsgBox lngId
    
    mvarobjProcNC.eliminar_equipamiento_implicado lngid
    PresentarDatos_EquipamientoImplicado

End Sub

Private Sub cmdEliminarPersonal_Click()

If lstPersonal.ListItems.Count = 0 Then Exit Sub

Dim lngid As Long

    lngid = lstPersonal.selectedItem
    'MsgBox lngId
    
    mvarobjProcNC.eliminar_personal_implicado lngid
    PresentarDatos_PersonalImplicado
    
End Sub

Private Sub cmdEliminarSecuencia_Click()
If lstSecuencia.ListItems.Count = 0 Then Exit Sub

Dim lngid As Long

    lngid = lstSecuencia.selectedItem
    'MsgBox lngId
    
    mvarobjProcNC.eliminar_secuencia_investigacion lngid
    PresentarDatos_Secuencia

End Sub

Private Sub cmdcancel_Click()

    If Not mvarblnEditable Then Unload Me

    If Not guardar_datos Then Exit Sub
    
    Unload Me
End Sub


Private Sub cmdSubirPrioridad_Click()

If lstSecuencia.ListItems.Count = 0 Then Exit Sub

Dim lngid As Long

    lngid = lstSecuencia.selectedItem
    'MsgBox lngId
    
    mvarobjProcNC.secuencia_cambiar_orden lngid, lngid - 1
    PresentarDatos_Secuencia
End Sub

Private Sub Form_Activate()
    Me.Move (Screen.Width - Me.Width) / 2, (Screen.Height - Me.Height) / 2

End Sub

Private Sub Form_Load()

    cabecera
    cargar_botones Me
    
    cargar_listados

    cargar_datos

    opciones_edicion

End Sub


Private Sub cargar_datos()

mvarobjProcNC.Carga PK
    
PresentarDatos_PersonalImplicado
PresentarDatos_EquipamientoImplicado
PresentarDatos_Secuencia
PresentarDatos_DocumentosAdjuntos

    With mvarobjProcNC
        txtRecoleccionDatos.Text = .getRECOLECCION_DATOS
        If .getANALISIS_ESCENA_FECHA = 0 Then
            txtFechaHora.value = Now
        Else
            txtFechaHora.value = CDate(Format(.getANALISIS_ESCENA_FECHA, "dd/mm/yyyy") & " " & Format(CStr(.getANALISIS_ESCENA_HORA) & ":" & CStr(.getANALISIS_ESCENA_MINUTOS), "Hh:Nn"))
            
        End If
        txtLocalizacion.Text = .getANALISIS_ESCENA_LOCALIZACION
        txtCondicionesAmbientales.Text = .getANALISIS_ESCENA_CONDICIONES_AMBIENTALES
        txtCondicionesOperacion.Text = .getANALISIS_ESCENA_CONDICIONES_OPERACION
        txtComunicacion.Text = .getANALISIS_ESCENA_COMUNICACION
        txtCambiosRecientes.Text = .getANALISIS_ESCENA_CAMBIOS_RECIENTES
        txtFormacion.Text = .getANALISIS_ESCENA_FORMACION
        txtExperiencia.Text = .getANALISIS_ESCENA_EXPERIENCIA
    End With
End Sub

Private Sub cabecera()
    With lstPersonal.ColumnHeaders
        .Add , , "id", 0, lvwColumnLeft
        .Add , , "Usuarios", lstPersonal.Width, lvwColumnLeft
    End With
        
    With lstEquipos.ColumnHeaders
        .Add , , "id", 0, lvwColumnLeft
        .Add , , "Equipos", lstEquipos.Width, lvwColumnLeft
    End With
    
    With lstSecuencia.ColumnHeaders
        .Add , , "orden", 0, lvwColumnLeft
        .Add , , "Secuencia Eventos", lstSecuencia.Width, lvwColumnLeft
    End With
    
    With lstDocumentacion.ColumnHeaders
        .Add , , "id", 0, lvwColumnLeft
        .Add , , "Documento", lstDocumentacion.Width, lvwColumnLeft
    End With
    
End Sub

Private Sub cargar_listados()


    ' Ahora carga las Combos
    llenar_combo cmbPersonal, New clsUsuarios, 0, frmUsuarios, ""
    llenar_combo cmbEquipos, New clsEquipos, 0, frmEquipoEdicion, ""
    
End Sub

Private Sub PresentarDatos_PersonalImplicado()
    ' Personal  Implicado
    lstPersonal.ListItems.Clear
    Set RS = mvarobjProcNC.devolver_listado_personal_implicado
    If RS.RecordCount <> 0 Then
        RS.MoveFirst
        While Not RS.EOF
            With lstPersonal.ListItems.Add(, , RS("id_usuario"))
                .SubItems(1) = RS("usuario")
            End With
            RS.MoveNext
        Wend
    End If
End Sub

Private Sub PresentarDatos_EquipamientoImplicado()
    ' Equipos Implicados
    lstEquipos.ListItems.Clear
    Set RS = mvarobjProcNC.devolver_listado_equipos_implicados
    If RS.RecordCount <> 0 Then
        RS.MoveFirst
        While Not RS.EOF
            With lstEquipos.ListItems.Add(, , RS("id_equipo"))
                .SubItems(1) = CStr(RS("equipo"))
            End With
            RS.MoveNext
        Wend
    End If
End Sub

Private Sub PresentarDatos_Secuencia()
    ' Secuencia de Eventos
    lstSecuencia.ListItems.Clear
    Set RS = mvarobjProcNC.devolver_listado_secuencia_investigacion
    If RS.RecordCount <> 0 Then
        RS.MoveFirst
        While Not RS.EOF
            With lstSecuencia.ListItems.Add(, , RS("orden"))
                .SubItems(1) = RS("descripcion")
            End With
            RS.MoveNext
        Wend
    End If
End Sub

Private Sub PresentarDatos_DocumentosAdjuntos()
    lstDocumentacion.ListItems.Clear
    Dim oAdjunto As New clsAdjuntos
    Dim RS As ADODB.RecordSet
    Set RS = oAdjunto.Listado(TOBJETO.TOBJETO_PROCNC_RECOLECCION_DATOS, PK, "", "")
    If RS.RecordCount > 0 Then
        Do
            With lstDocumentacion.ListItems.Add(, , RS(0))
                 .SubItems(1) = RS(2)
            End With
            RS.MoveNext
        Loop Until RS.EOF
    End If
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
If UnloadMode = 0 Then cmdcancel_Click
End Sub

Private Sub lstDocumentacion_DblClick()
    If lstDocumentacion.ListItems.Count = 0 Then Exit Sub
    Dim oAdjunto As New clsAdjuntos
    oAdjunto.CargarDocumento TOBJETO.TOBJETO_PROCNC_RECOLECCION_DATOS, PK, 0, lstDocumentacion.ListItems(lstDocumentacion.selectedItem.Index).Text, True
    Set oAdjunto = Nothing
End Sub
Public Property Get Editable() As Boolean
    Editable = mvarblnEditable
End Property

Public Property Let Editable(ByVal blnEditable As Boolean)
    mvarblnEditable = blnEditable
End Property
