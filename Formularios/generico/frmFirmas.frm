VERSION 5.00
Begin VB.Form frmFirmas 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Firma"
   ClientHeight    =   4290
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7365
   ClipControls    =   0   'False
   HasDC           =   0   'False
   Icon            =   "frmFirmas.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4290
   ScaleWidth      =   7365
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdDocumento 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Ver Documento"
      Height          =   1005
      Left            =   90
      Picture         =   "frmFirmas.frx":6852
      Style           =   1  'Graphical
      TabIndex        =   17
      Top             =   3240
      Visible         =   0   'False
      Width           =   1320
   End
   Begin VB.CommandButton cmdDocCurso 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Ver Doc. Curso"
      Height          =   1005
      Left            =   90
      Picture         =   "frmFirmas.frx":711C
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   3240
      Visible         =   0   'False
      Width           =   1320
   End
   Begin VB.TextBox txtID 
      Height          =   285
      Left            =   3555
      Locked          =   -1  'True
      TabIndex        =   14
      Top             =   3735
      Visible         =   0   'False
      Width           =   825
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0C0C0&
      Height          =   2445
      Left            =   45
      TabIndex        =   2
      Top             =   720
      Width           =   7260
      Begin VB.TextBox txtFechaF 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         Height          =   285
         Left            =   5535
         Locked          =   -1  'True
         TabIndex        =   11
         Top             =   1080
         Width           =   1455
      End
      Begin VB.TextBox txtTipo 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000001&
         Height          =   285
         Left            =   1755
         Locked          =   -1  'True
         TabIndex        =   10
         Top             =   270
         Width           =   5235
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
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
         Left            =   1755
         Locked          =   -1  'True
         TabIndex        =   9
         Top             =   675
         Width           =   5235
      End
      Begin VB.TextBox txtFechaI 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         Height          =   285
         Left            =   1755
         Locked          =   -1  'True
         TabIndex        =   8
         Top             =   1080
         Width           =   1455
      End
      Begin VB.TextBox txtDescripcion 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         Height          =   870
         Left            =   1755
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   7
         Top             =   1485
         Width           =   5235
      End
      Begin VB.Label lblFechaF 
         BackColor       =   &H00C0C0C0&
         Caption         =   "FECHA:"
         Height          =   195
         Left            =   4005
         TabIndex        =   12
         Top             =   1125
         Width           =   1905
      End
      Begin VB.Label Label5 
         BackColor       =   &H00C0C0C0&
         Caption         =   "DESCRIPCIÓN:"
         Height          =   195
         Left            =   180
         TabIndex        =   6
         Top             =   1485
         Width           =   1140
      End
      Begin VB.Label lblFechaI 
         BackColor       =   &H00C0C0C0&
         Caption         =   "FECHA:"
         Height          =   195
         Left            =   180
         TabIndex        =   5
         Top             =   1080
         Width           =   1635
      End
      Begin VB.Label lblCodigo 
         BackColor       =   &H00C0C0C0&
         Caption         =   "CÓDIGO:"
         Height          =   195
         Left            =   180
         TabIndex        =   4
         Top             =   675
         Width           =   1500
      End
      Begin VB.Label Label2 
         BackColor       =   &H00C0C0C0&
         Caption         =   "¿QUÉ FIRMAS?"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00008000&
         Height          =   195
         Left            =   180
         TabIndex        =   3
         Top             =   315
         Width           =   1500
      End
   End
   Begin VB.CommandButton cmdSalir 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Cancelar"
      Height          =   1005
      Left            =   5895
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   3240
      Width           =   1365
   End
   Begin VB.CommandButton cmdAceptar 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Firmar"
      Height          =   1005
      Left            =   4500
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   3240
      Width           =   1365
   End
   Begin VB.Label lblnombre 
      Alignment       =   2  'Center
      BackColor       =   &H80000009&
      Caption         =   " "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   285
      Left            =   315
      TabIndex        =   18
      Top             =   90
      Width           =   6585
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0C0C0&
      Caption         =   "ID FIRMA:"
      Height          =   195
      Left            =   2565
      TabIndex        =   15
      Top             =   3780
      Visible         =   0   'False
      Width           =   780
   End
   Begin VB.Label lblSuperior 
      Alignment       =   2  'Center
      BackColor       =   &H80000009&
      Caption         =   "La firma digital tendrá la misma validez que la firma convencional"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   90
      TabIndex        =   13
      Top             =   450
      Width           =   7035
   End
   Begin VB.Shape fondo 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   735
      Left            =   -360
      Top             =   0
      Width           =   7695
   End
End
Attribute VB_Name = "frmFirmas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public ID_FIRMA As Long
Private oFirma As New clsFirmas
Private oCurso As New clsFormacion_cursos
Private TOBJETO As Long

Private Sub cmdAceptar_Click()
    On Error GoTo firmar_Click
    oFirma.setFIRMADA = 1
    oFirma.Modificar
    Select Case TOBJETO
        Case TOBJETO_ASISTENCIA_CURSO_ASISTENTES
            frmFormacion_Evaluacion.PK = oCurso.getID_CURSO
            frmFormacion_Evaluacion.ID_ASISTENTE = oFirma.getUSUARIO_ID
            frmFormacion_Evaluacion.Show 1
        Case TOBJETO_DOCUMENTO_VALORACION
            frmCA_Valoracion.PK_DOCUMENTO_ID = oFirma.getCOBJETO
            frmCA_Valoracion.PK_USUARIO_ID = 0
            frmCA_Valoracion.Show 1
        Case other
            MsgBox "Se ha registrado la firma con éxito", vbOKOnly + vbInformation, App.Title
    End Select
    frmTelefonos.cargar_lista_firmas
'    CUALIFICACION
    Unload Me
    Exit Sub
firmar_Click:
    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure Firmar_Click of Formulario frmFirmas"
End Sub
Private Sub cmdCancelar_Click()
   Unload Me
End Sub
Private Sub cmdDocCurso_Click()
    Dim strCad As String
    Dim arrNom() As String
    Dim arrVal() As String
    Dim objfrm As New frmReport
    With objfrm
        .iniciar
        .informe = "Formacion\rptCurso"
        .ParametrosNombre = arrNom
        .ParametrosValores = arrVal
        .criterio = "{formacion_cursos.ID_CURSO} = " & CLng(oFirma.getCOBJETO)
        .imprimir = False
        .generar
        .Show 1
    End With
End Sub
Private Sub cmdDocumento_Click()
    Dim strCad As String
    Dim arrNom() As String
    Dim arrVal() As String
    Dim objfrm As New frmReport
    Select Case TOBJETO
        Case TOBJETO_ASISTENCIA_CURSO_ASISTENTES To TOBJETO_ASISTENCIA_CURSO_CALIDAD
            With objfrm
                .iniciar
                .informe = "Formacion\rptCurso"
                .ParametrosNombre = arrNom
                .ParametrosValores = arrVal
                .criterio = "{formacion_cursos.ID_CURSO} = " & CLng(oFirma.getCOBJETO)
                .imprimir = False
                .generar
                .Show 1
            End With
        Case TOBJETO_CERTIFICACION To TOBJETO_CERTIFICACION_RESPONSABLE
            With objfrm
                .iniciar
                .informe = "Formacion\rptCertificacion"
                .ParametrosNombre = arrNom
                .ParametrosValores = arrVal
                .criterio = "{formacion_certificados.ID_FORMACION_CERTIFICADO} = " & CLng(oFirma.getCOBJETO)
                .imprimir = False
                .generar
                .Show 1
            End With
        Case TOBJETO_DOCUMENTO_VALORACION
            Dim oCA_Documento As New clsCa_documentos
            oCA_Documento.mostrar oFirma.getCOBJETO, True
            Set oCA_Documento = Nothing
    End Select
End Sub
Private Sub cmdSalir_Click()
    Unload Me
End Sub
Private Sub Form_Load()
    log (Me.Name)
    cargar_botones Me
    carga_formulario
End Sub
Private Sub carga_formulario()
    'Carga de la firma a través de su ID
    'Segun el TOBJETO se recuperarán las descripciones de un sitio u otro
    'De momento sólo se habilita el detalle para la firma de cursos
    Dim oempleado As New clsEmpleados
    Dim rs As New ADODB.Recordset
    
    oFirma.Carga (ID_FIRMA)
    oempleado.CARGAR oFirma.getUSUARIO_ID
    lblnombre.Caption = oempleado.getNOMBRE
    evalua_tobjeto

    oempleado.CARGAR_POR_USUARIO USUARIO.getID_EMPLEADO
    If oFirma.getUSUARIO_ID = oempleado.getID_EMPLEADO Or UCase(USUARIO.getUSUARIO) = "JULIO" Or UCase(USUARIO.getUSUARIO) = "MARIBEL" Then
        cmdAceptar.Enabled = True
    Else
        cmdAceptar.Enabled = False
    End If
    Set oempleado = Nothing

End Sub

Private Sub evalua_tobjeto()

    TOBJETO = oFirma.getTOBJETO
    cmdDocCurso.Visible = False
    cmdDocumento.Visible = True
    Select Case TOBJETO
        Case TOBJETO_ASISTENCIA_CURSO_ASISTENTES To TOBJETO_ASISTENCIA_CURSO_CALIDAD
            detalle_curso
        Case TOBJETO_INVITACION_CURSO
        Case TOBJETO_CERTIFICACION To TOBJETO_CERTIFICACION_RESPONSABLE
            detalle_certificacion
        Case TOBJETO_DOCUMENTO_VALORACION
            detalle_documento_valoracion
        Case other
            detalle_otros
    End Select
  
End Sub
Private Sub detalle_documento_valoracion()
    txtID.Text = ID_FIRMA
    Dim oDecodificadora As New clsDecodificadora
    oDecodificadora.Carga_valor PARAM_TOBJETO, TOBJETO
    lblCodigo.Caption = "Documento:"
    txtTipo.Text = oDecodificadora.getDESCRIPCION
    Set oDecodificadora = Nothing
    Dim oca_doc As New clsCa_documentos
    oca_doc.Carga oFirma.getCOBJETO
    txtCodigo.Text = oca_doc.getNOMBRE
    txtdescripcion.Text = "Una vez leído el documento, debe rellenar un cuestionario para evaluar el procedimiento"
    lblFechaI.Caption = "Fecha :"
    txtFechaI.Text = Format(oFirma.getFTIMESTP, "dd/mm/yyyy")
    lblFechaF.Visible = False
    txtFechaF.Visible = False
End Sub
Private Sub detalle_curso()
    On Error GoTo error_detalle_curso
    oCurso.Carga (oFirma.getCOBJETO)
    txtCodigo.Text = "RFI-" & Format(oCurso.getCOD_CURSO, "000") & "/" & Year(Date)
    txtID.Text = ID_FIRMA
    Dim oDecodificadora As New clsDecodificadora
     
    oDecodificadora.Carga_valor PARAM_TOBJETO, TOBJETO
    lblCodigo.Caption = "CÓDIGO CURSO:"
    txtTipo.Text = oDecodificadora.getDESCRIPCION
    
    Set oDecodificadora = Nothing
     
    If TOBJETO >= TOBJETO_ASISTENCIA_CURSO_ASISTENTES And TOBJETO <= TOBJETO_ASISTENCIA_CURSO_CALIDAD Then
        lblFechaI.Caption = "INICIO CURSO:"
        lblFechaF.Caption = "FIN CURSO:"

        txtFechaI.Text = oCurso.getFECHA_REAL_I
        txtFechaF.Text = oCurso.getFECHA_REAL_F
    Else
        lblFechaI.Caption = "INICIO PREVISTO:"
        lblFechaF.Caption = "FIN PREVISTO:"

        txtFechaI.Text = oCurso.getFECHA_PREVISTA_I
        txtFechaF.Text = oCurso.getFECHA_PREVISTA_F
    End If
 
    txtdescripcion.Text = oCurso.getDESCRIPCION
    
    Exit Sub
    
error_detalle_curso:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure detalle_curso of Formulario frmFirmas"

End Sub
'M1143-I
Private Sub detalle_certificacion()
    'Detalles del curso que va a firmarse
    'El campo COBJETO de la firma contendrá la clave de la certificación en caso de TOBJETO = 25
    
    Dim oCertificado As New clsFormacion_certificados
    Dim oDOCUMENTO As New clsCa_documentos
    
    On Error GoTo error_detalle_certificacion
    
    oCertificado.Carga (oFirma.getCOBJETO)
    txtCodigo.Text = "CF " & Format(oCertificado.getCOD_CERTIFICADO, "000") & "-" & Format(oCertificado.getANYO, "0000")
    txtID.Text = ID_FIRMA

    Dim oDecodificadora As New clsDecodificadora
     
    oDecodificadora.Carga_valor PARAM_TOBJETO, TOBJETO
    lblCodigo.Caption = "CÓDIGO CERTF.:"
    txtTipo.Text = oDecodificadora.getDESCRIPCION
     
    Set oDecodificadora = Nothing

    lblFechaI.Caption = "FECHA CERTIFICADO:"
    lblFechaF.Caption = " "
    txtFechaI.Text = Format(oCertificado.getFECHA_CERTIFICACION, "yyyy-mm-dd")
    txtFechaF.Visible = False
    oDOCUMENTO.Carga oCertificado.getDOCUMENTO_ID
    txtdescripcion.Text = "Certifica:" & oDOCUMENTO.getNOMBRE
    
    Exit Sub
    
error_detalle_certificacion:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure detalle_certificacion of Formulario frmFirmas"

End Sub
Private Sub detalle_otros()

End Sub

Private Sub CUALIFICACION()
    On Error GoTo cmdok_Click_Error
    'Carga de los datos del curso
      Dim oCurso As New clsFormacion_cursos
      Dim oFormador As New clsFormacion_Formadores
      Dim rsFormadores As New ADODB.Recordset
      oCurso.Carga oFirma.getCOBJETO
      Set rsFormadores = oFormador.Listado_internos(oFirma.getCOBJETO)
      
      'Matriz de cualificaciones
      '-------------------------
      ' Por cada documento relacionado con el curso
      ' se marcará la lista de asistentes completa
      ' Sólo se ejecuta si el curso está vinculado a un plan de formación (getPlan_id >0)
      
      If oCurso.getPLAN_ID > 0 Then
         Dim strMsg As String
         Dim rsDocumentos As New ADODB.Recordset                       'Recordset por los documentos del curso
         Dim oDocumentos As New clsFormacion_pf_docs       'Documentos del curso
         Set rsDocumentos = oDocumentos.Listado_Plan(oCurso.getPLAN_ID)
    
         If rsDocumentos.RecordCount > 0 Then
              strMsg = " Se ha acreditado en la formación teórica de los siguientes PNTs: " & vbCrLf
              strMsg = strMsg & "---------------------------------------------------------------------------------------------- " & vbCrLf
              Do
                   Dim oDetalle As New clsCa_documentos
                   Dim oCualificaciones As New clsEmpleados_cualificaciones
                   Dim rsCualificaciones As New ADODB.Recordset
                   
                   oDetalle.Carga rsDocumentos("DOCUMENTO_ID")
                   strMsg = strMsg & "(" & oDetalle.getCODIGO & ") " & oDetalle.getNOMBRE & vbCrLf
                   Set rsCualificaciones = oCualificaciones.Listado_Empleado_DOC(oFirma.getUSUARIO_ID, rsDocumentos("DOCUMENTO_ID"), 0)
                   If rsCualificaciones.RecordCount = 0 Then
                       With oCualificaciones
                           .setEMPLEADO_ID = oFirma.getUSUARIO_ID
                           .setDOCUMENTO_ID = rsDocumentos("DOCUMENTO_ID")
                           .setEMPLEADO_ID_FORMADOR = rsFormadores("ID_EMPLEADO")
                           .setEN_HISTORICO = 0
                           .setES_FORMADOR = 0
                           .setESTADO = 0
                           .setFECHA_FIRMA_DIRECTOR = Format(oCurso.getFECHA_REAL_F, "yyyy-mm-dd")
                           .setFECHA_FIRMA_FORMADOR = Format(oCurso.getFECHA_REAL_F, "yyyy-mm-dd")
                           .setFECHA_FIRMA_TECNICO = Format(oCurso.getFECHA_REAL_F, "yyyy-mm-dd")
                           .setFECHA_FORMACION_TEORICA = Format(oCurso.getFECHA_REAL_F, "yyyy-mm-dd")
                           .setFECHA_ULTIMA_RECUALIFICACION = "1900-01-01"
                           .setFORMADOR_NO_CUALIFICADO = 0
                           .setID_CUALIFICACION = 0
                           .setTEXTO_FORMACION_TEORICA = "Lectura del PNT y explicación por parte del formador."
                           .Insertar
                       End With
                   End If
                           
                   Set rsCualificaciones = Nothing
                   Set oDetalle = Nothing
                   Set oCualificaciones = Nothing
                   
                   rsDocumentos.MoveNext
              Loop Until rsDocumentos.EOF
              
              MsgBox strMsg, vbInformation + vbOKOnly, App.Title
        Else
              MsgBox "El plan de formación no tiene PNTs sobre los que cualificarse", vbInformation + vbOKOnly, App.Title
        End If
    End If
    Set rsDocumentos = Nothing

    Exit Sub
cmdok_Click_Error:
    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure cmdCualificar of Formulario frmFormacion_Curso"
End Sub
