VERSION 5.00
Begin VB.Form frmFormacion_CF_Detalle 
   BackColor       =   &H00C0C0C0&
   Caption         =   "Certificaciones"
   ClientHeight    =   9300
   ClientLeft      =   6240
   ClientTop       =   1095
   ClientWidth     =   8175
   LinkTopic       =   "Form1"
   ScaleHeight     =   9300
   ScaleWidth      =   8175
   Begin VB.Frame Frame3 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Documentación a certificar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   690
      Left            =   0
      TabIndex        =   18
      Top             =   3780
      Width           =   8160
      Begin VB.TextBox txtDocumentacion 
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   135
         MultiLine       =   -1  'True
         TabIndex        =   20
         Top             =   270
         Width           =   7935
      End
   End
   Begin VB.Frame frmFormacion_CF_Detalle 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Empleado certificado"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   690
      Left            =   0
      TabIndex        =   17
      Top             =   1620
      Width           =   8160
      Begin VB.TextBox txtEmpleado 
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   135
         MultiLine       =   -1  'True
         TabIndex        =   21
         Top             =   270
         Width           =   7935
      End
   End
   Begin VB.CommandButton cmdImprimir 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Certificado"
      Height          =   825
      Left            =   4680
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   8460
      Width           =   1230
   End
   Begin VB.Frame Frame6 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Certificación"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   0
      TabIndex        =   9
      Top             =   855
      Width           =   8160
      Begin VB.TextBox txtAnyo 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   5265
         Locked          =   -1  'True
         TabIndex        =   12
         Top             =   270
         Width           =   1320
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000001&
         Height          =   330
         Left            =   1755
         Locked          =   -1  'True
         TabIndex        =   10
         Top             =   270
         Width           =   1770
      End
      Begin VB.Label Label3 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Año"
         Height          =   240
         Left            =   4770
         TabIndex        =   13
         Top             =   315
         Width           =   375
      End
      Begin VB.Label Label2 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Código"
         Height          =   240
         Left            =   1080
         TabIndex        =   11
         Top             =   315
         Width           =   555
      End
   End
   Begin VB.Frame Frame5 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Detalle de la certificación"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3885
      Left            =   0
      TabIndex        =   4
      Top             =   4500
      Width           =   8160
      Begin VB.CommandButton cmdAnadir 
         BackColor       =   &H00E0E0E0&
         Height          =   510
         Left            =   7380
         Style           =   1  'Graphical
         TabIndex        =   24
         Top             =   3330
         Width           =   690
      End
      Begin VB.ComboBox cmbBase 
         Height          =   315
         Left            =   2430
         TabIndex        =   22
         Top             =   3375
         Width           =   4830
      End
      Begin VB.TextBox txtBase 
         Appearance      =   0  'Flat
         Height          =   1950
         Left            =   90
         MultiLine       =   -1  'True
         TabIndex        =   7
         Top             =   1350
         Width           =   7980
      End
      Begin VB.TextBox txtMateria 
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   90
         MultiLine       =   -1  'True
         TabIndex        =   5
         Top             =   585
         Width           =   7935
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0C0C0&
         Caption         =   "Listado de cualidades:"
         Height          =   240
         Left            =   180
         TabIndex        =   23
         Top             =   3420
         Width           =   2040
      End
      Begin VB.Label Label1 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Base"
         Height          =   240
         Left            =   180
         TabIndex        =   8
         Top             =   1035
         Width           =   1320
      End
      Begin VB.Label Label19 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Materia"
         Height          =   240
         Left            =   180
         TabIndex        =   6
         Top             =   315
         Width           =   1185
      End
   End
   Begin VB.Frame cmbResponsable 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Responsable de departamento"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   690
      Left            =   0
      TabIndex        =   3
      Top             =   3060
      Width           =   8160
      Begin VB.TextBox txtResponsable 
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   135
         MultiLine       =   -1  'True
         TabIndex        =   25
         Top             =   270
         Width           =   7935
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Certificador de personal"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   690
      Left            =   0
      TabIndex        =   2
      Top             =   2340
      Width           =   8160
      Begin VB.TextBox txtCertificador 
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   135
         MultiLine       =   -1  'True
         TabIndex        =   19
         Top             =   270
         Width           =   7935
      End
   End
   Begin VB.CommandButton cmdOk 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Aceptar"
      Height          =   825
      Left            =   5985
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   8460
      Width           =   1050
   End
   Begin VB.CommandButton cmdcancel 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Salir"
      Height          =   825
      Left            =   7110
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   8460
      Width           =   1050
   End
   Begin VB.Label lbltitulo 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Certificado de Formador"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000001&
      Height          =   330
      Left            =   990
      TabIndex        =   15
      Top             =   90
      Width           =   5805
   End
   Begin VB.Label lblSubtitulo 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Seleccione la formación y el empleado que será certificado"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000006&
      Height          =   195
      Left            =   1035
      TabIndex        =   14
      Top             =   495
      Width           =   5805
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   7335
      Picture         =   "frmFormacion_CF_Detalle.frx":0000
      Top             =   135
      Width           =   480
   End
   Begin VB.Shape fondo 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   825
      Left            =   -495
      Top             =   0
      Width           =   8685
   End
End
Attribute VB_Name = "frmFormacion_CF_Detalle"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public PK As Long
Public CUALIFICACION As Long
Public ID_DOC As Long

Private Sub Form_Load()
    log Me.Name
    cargar_botones Me
    cargar_combos
    If CUALIFICACION <> 0 Then
       obtenerPK
       cmdok.Enabled = True
    Else
       cmdok.Enabled = False
    End If
    If CUALIFICACION <> 0 And PK = 0 Then
       cmdImprimir.Enabled = False
       formularioAlta
    Else
       cmdImprimir.Enabled = True
       formularioConsulta
    End If
End Sub

Private Sub obtenerPK()
    Dim oCertificado As New clsFormacion_certificados
    PK = oCertificado.CargaCualificacion(CUALIFICACION)
End Sub
Private Sub formularioConsulta()
On Error GoTo err_Consulta

    Dim oCertificado As New clsFormacion_certificados
    Dim oempleado As New clsEmpleados
    Dim oDoc As New clsCa_documentos
    Dim oCualificacion As New clsEmpleados_cualificaciones
    
    oCertificado.Carga PK
    oCualificacion.Carga oCertificado.getCUALIFICACION_ID
    txtCodigo = "CF " & oCertificado.getCOD_CERTIFICADO & " - " & oCertificado.getANYO
    txtAnyo.Text = oCertificado.getANYO
    txtAnyo.Visible = False
    Label3.Visible = False
    txtBase = oCertificado.getBASE
    oempleado.CARGAR oCualificacion.getEMPLEADO_ID
    txtEmpleado = oempleado.getNOMBRE
    oempleado.CARGAR oCertificado.getCERTIFICADOR_ID
    txtCertificador = oempleado.getNOMBRE
    oempleado.CARGAR oCualificacion.getEMPLEADO_ID_FORMADOR
    txtResponsable = oempleado.getNOMBRE
    oDoc.Carga oCertificado.getDOCUMENTO_ID
    txtDocumentacion = oDoc.getCODIGO
    txtMateria = oDoc.getNOMBRE
Exit Sub
err_Consulta:
 MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure consulta of Formulario frmFormacion_CF_Detalle"
  
End Sub

Private Sub formularioAlta()
    Dim oempleado As New clsEmpleados
    Dim oPNT As New clsCa_documentos
    Dim oCualificacion As New clsEmpleados_cualificaciones
    oCualificacion.Carga CUALIFICACION
    oempleado.CARGAR_POR_USUARIO (USUARIO.getID_EMPLEADO)
    txtCertificador.Text = oempleado.getNOMBRE
    oPNT.Carga ID_DOC
    oempleado.CARGAR oCualificacion.getEMPLEADO_ID
    txtEmpleado.Text = oempleado.getNOMBRE
    oempleado.CARGAR oCualificacion.getEMPLEADO_ID_FORMADOR
    txtResponsable.Text = oempleado.getNOMBRE
    txtDocumentacion.Text = oPNT.getCODIGO
    txtMateria.Text = oPNT.getNOMBRE
    txtCodigo.Text = "--"
    txtAnyo.Text = Format(Date, "yyyy")
    
    Set oempleado = Nothing
    Set oPNT = Nothing
End Sub

Private Sub anadirCertificado()
    On Error GoTo ErrAnadirCertificado
    Dim strMsg As String
    If PK = 0 Then
        strMsg = "Se va a generar un certificado de formación. ¿Desea continuar?"
    Else
        strMsg = "Se va a regenerar el certificado de formación. ¿Desea continuar?"
    End If
    
    If MsgBox(strMsg, vbQuestion + vbYesNo, App.Title) = vbYes Then
        Dim oCertificado As New clsFormacion_certificados
        Dim Pos As Integer
        Dim oempleado As New clsEmpleados
        oempleado.CARGAR_POR_USUARIO USUARIO.getID_EMPLEADO
    
        With oCertificado
         .EliminarCualificacion CUALIFICACION
         .setANYO = txtAnyo.Text
         .setBASE = txtBase.Text
         .setCUALIFICACION_ID = CUALIFICACION
         .setCERTIFICADOR_ID = oempleado.getID_EMPLEADO
         .setDOCUMENTO_ID = ID_DOC
         .setCUSERID = USUARIO.getID_EMPLEADO
         .setFECHA_CERTIFICACION = Format(Date, "yyyy-mm-dd")
         PK = .Insertar
         'Los registros se insertan firmados para evitar inconsistencia en los datos. No genera aviso al usuario.
         .generar_firmas PK
        End With
        
        MsgBox "La certificación ha sido registrada correctamente ", vbOKOnly
        generaDocumento ("\TEMP\CERTIFICADO_FORMACION.pdf")
        cmdImprimir.Enabled = True
        Unload Me
    End If
    Exit Sub
ErrAnadirCertificado:
    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure anadirCertificado of Formulario frmFormacion_CF_Detalle"
    
End Sub

Private Sub cmdcancel_Click()
    documento_escaner = ""
    Unload Me
End Sub

Private Sub cargar_combos()
    cargarCualidades
End Sub
Private Sub cargarCualidades()
    With cmbBase
        .AddItem "Formación técnica cualificada", 0
        .AddItem "Capacidad didáctica", 1
        .AddItem "Certificación de AiM", 2
        .AddItem "Puesta en marcha del método", 3
        .AddItem "Estudio en profundidad de la normativa", 4
        .AddItem "Validación del método de ensayo", 5
        .AddItem "Estudio en profundidad de este ensayo", 6
        .AddItem "Capacidad técnica", 7
        .AddItem "Revisión del método", 8
        .AddItem "Trayectoria profesional", 9
        .AddItem "Experiencia profesional", 10
    End With
End Sub

Private Sub generaDocumento(pdf As String)
    Dim strCad As String
    Dim arrNom() As String
    Dim arrVal() As String
    Dim objfrm As New frmReport
    Dim Path As String
    Path = ReadINI(App.Path + "\config.ini", "Documentos", "ca_evidencias")
On Error Resume Next
    MkDir Path & "\TEMP"
On Error GoTo ERROR
    With objfrm
        .iniciar
        .informe = "Formacion\rptCertificacion"
        If Trim(pdf) <> "" Then
            .pdf = Path & pdf
        Else
            .pdf = ""
        End If
        documento_escaner = Trim(.pdf)
        .ParametrosNombre = arrNom
        .ParametrosValores = arrVal
        .criterio = "{formacion_certificados.ID_FORMACION_CERTIFICADO} = " & PK & " and {firmas_1.TOBJETO} = 27 AND {firmas_2.TOBJETO} = 26 AND {firmas.TOBJETO} = 25"
        .imprimir = True
        .generar
        .Show 1
    End With
    Exit Sub
ERROR:
    documento_escaner = ""
    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure generaDocumento of Formulario frmFormacion_CF_Detalle"
    Exit Sub
End Sub

Private Sub cmdImprimir_Click()
    generaDocumento ("")
End Sub
Private Sub cmdok_Click()
    anadirCertificado
End Sub

Private Sub cmdAnadir_Click()
    txtBase = txtBase + vbCrLf + " - " + cmbBase.Text + vbCrLf
    cmbBase.Text = ""
End Sub
