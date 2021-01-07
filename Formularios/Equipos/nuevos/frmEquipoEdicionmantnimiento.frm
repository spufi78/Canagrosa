VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{CEA3B7F6-2847-4E5E-A551-DB7A62489D44}#34.0#0"; "miCombo.ocx"
Begin VB.Form frmEquipoEdicionmantnimiento 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   6270
   ClientLeft      =   3075
   ClientTop       =   3615
   ClientWidth     =   10170
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "frmEquipoEdicionmantnimiento.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6270
   ScaleWidth      =   10170
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdEtiqueta 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Etiqueta"
      Height          =   900
      Left            =   90
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   5310
      Width           =   1050
   End
   Begin MSComDlg.CommonDialog cd 
      Left            =   3930
      Top             =   5640
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Frame marcoDatos_Verificacion 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   "Datos de la verificación"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   1515
      Left            =   45
      TabIndex        =   9
      Top             =   3750
      Width           =   10080
      Begin VB.CheckBox chkConforme 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Conforme"
         Height          =   240
         Left            =   8940
         TabIndex        =   17
         Top             =   270
         Width           =   1080
      End
      Begin VB.TextBox txtDatos 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   750
         Index           =   5
         Left            =   1620
         MaxLength       =   100
         TabIndex        =   4
         Top             =   630
         Width           =   8340
      End
      Begin pryCombo.miCombo cmbOperador_Interno_Real 
         Height          =   330
         Left            =   1620
         TabIndex        =   3
         Top             =   270
         Width           =   5475
         _ExtentX        =   9657
         _ExtentY        =   582
      End
      Begin pryCombo.miCombo cmbOperador_Externo_Real 
         Height          =   330
         Left            =   1620
         TabIndex        =   12
         Top             =   270
         Width           =   5475
         _ExtentX        =   9657
         _ExtentY        =   582
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Realizado por"
         Height          =   195
         Index           =   6
         Left            =   135
         TabIndex        =   11
         Top             =   315
         Width           =   975
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Limitaciones uso"
         Height          =   195
         Index           =   5
         Left            =   135
         TabIndex        =   10
         Top             =   675
         Width           =   1170
      End
   End
   Begin VB.CommandButton cmdcancel 
      BackColor       =   &H00E0E0E0&
      Caption         =   "ESC-Salir"
      Height          =   870
      Left            =   9090
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   5325
      Width           =   1050
   End
   Begin VB.CommandButton cmdok 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Aceptar"
      Height          =   870
      Left            =   8010
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   5325
      Width           =   1050
   End
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   "Verificación"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   3165
      Left            =   45
      TabIndex        =   7
      Top             =   510
      Width           =   10080
      Begin VB.TextBox txtDatos 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   330
         Index           =   3
         Left            =   1605
         Locked          =   -1  'True
         MaxLength       =   255
         TabIndex        =   25
         Top             =   2295
         Width           =   6540
      End
      Begin VB.CommandButton cmdAbrirDocumento 
         BackColor       =   &H00E0E0E0&
         Height          =   420
         Left            =   9345
         Picture         =   "frmEquipoEdicionmantnimiento.frx":000C
         Style           =   1  'Graphical
         TabIndex        =   24
         ToolTipText     =   "Ver norma"
         Top             =   2250
         Width           =   465
      End
      Begin VB.CommandButton cmdExplorarDocumento 
         BackColor       =   &H00E0E0E0&
         Height          =   420
         Left            =   8310
         Picture         =   "frmEquipoEdicionmantnimiento.frx":0261
         Style           =   1  'Graphical
         TabIndex        =   23
         ToolTipText     =   "Buscar documento"
         Top             =   2250
         Width           =   465
      End
      Begin VB.CommandButton cmdEliminarDocumento 
         BackColor       =   &H00E0E0E0&
         Height          =   420
         Left            =   8850
         Picture         =   "frmEquipoEdicionmantnimiento.frx":04D2
         Style           =   1  'Graphical
         TabIndex        =   22
         ToolTipText     =   "Eliminar documento"
         Top             =   2250
         Width           =   420
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H00E0E0E0&
         Height          =   420
         Left            =   8850
         Picture         =   "frmEquipoEdicionmantnimiento.frx":0666
         Style           =   1  'Graphical
         TabIndex        =   21
         ToolTipText     =   "Eliminar documento"
         Top             =   1770
         Width           =   420
      End
      Begin VB.CommandButton Command2 
         BackColor       =   &H00E0E0E0&
         Height          =   420
         Left            =   8310
         Picture         =   "frmEquipoEdicionmantnimiento.frx":07FA
         Style           =   1  'Graphical
         TabIndex        =   20
         ToolTipText     =   "Buscar documento"
         Top             =   1770
         Width           =   465
      End
      Begin VB.CommandButton Command3 
         BackColor       =   &H00E0E0E0&
         Height          =   420
         Left            =   9345
         Picture         =   "frmEquipoEdicionmantnimiento.frx":0A6B
         Style           =   1  'Graphical
         TabIndex        =   19
         ToolTipText     =   "Ver norma"
         Top             =   1770
         Width           =   465
      End
      Begin VB.TextBox txtDatos 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   330
         Index           =   0
         Left            =   1605
         Locked          =   -1  'True
         MaxLength       =   255
         TabIndex        =   18
         Top             =   1815
         Width           =   6540
      End
      Begin VB.TextBox txtDatos 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   330
         Index           =   1
         Left            =   6135
         MaxLength       =   50
         TabIndex        =   0
         Top             =   2730
         Width           =   930
      End
      Begin VB.TextBox txtDatos 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   330
         Index           =   2
         Left            =   7215
         MaxLength       =   50
         TabIndex        =   1
         Top             =   2730
         Width           =   930
      End
      Begin pryCombo.miCombo cmbUnidad 
         Height          =   330
         Left            =   1605
         TabIndex        =   2
         Top             =   2730
         Width           =   2760
         _ExtentX        =   4868
         _ExtentY        =   582
      End
      Begin MSComCtl2.DTPicker fecha_actual 
         Height          =   345
         Left            =   8235
         TabIndex        =   26
         Top             =   210
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   609
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
         Format          =   70713345
         CurrentDate     =   2
         MinDate         =   2
      End
      Begin MSComCtl2.DTPicker fecha_proxima 
         Height          =   345
         Left            =   8235
         TabIndex        =   27
         Top             =   585
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   609
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
         Format          =   70713345
         CurrentDate     =   2
         MinDate         =   2
      End
      Begin MSDataListLib.DataCombo cmbTipoCalibracion 
         Height          =   315
         Left            =   1605
         TabIndex        =   28
         Top             =   210
         Width           =   4245
         _ExtentX        =   7488
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         Style           =   2
         Text            =   ""
      End
      Begin MSDataListLib.DataCombo cmbPeriCalibracion 
         Height          =   315
         Left            =   1605
         TabIndex        =   29
         Top             =   600
         Width           =   4260
         _ExtentX        =   7514
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         Style           =   2
         Text            =   ""
      End
      Begin pryCombo.miCombo cmbCalibrador 
         Height          =   330
         Left            =   1605
         TabIndex        =   30
         Top             =   990
         Width           =   6525
         _ExtentX        =   11509
         _ExtentY        =   582
      End
      Begin pryCombo.miCombo miCombo1 
         Height          =   330
         Left            =   1605
         TabIndex        =   31
         Top             =   1380
         Width           =   6525
         _ExtentX        =   11509
         _ExtentY        =   582
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Tipo"
         Height          =   195
         Index           =   3
         Left            =   135
         TabIndex        =   39
         Top             =   345
         Width           =   315
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Procedimiento"
         Height          =   195
         Index           =   4
         Left            =   135
         TabIndex        =   38
         Top             =   1425
         Width           =   1005
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Responsable"
         Height          =   195
         Index           =   8
         Left            =   135
         TabIndex        =   37
         Top             =   1035
         Width           =   930
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "F. Actual Verificación"
         Height          =   195
         Index           =   10
         Left            =   6645
         TabIndex        =   36
         Top             =   270
         Width           =   1500
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "F. Próx. Verificación"
         Height          =   195
         Index           =   0
         Left            =   6645
         TabIndex        =   35
         Top             =   630
         Width           =   1410
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Periodo Calibración"
         Height          =   195
         Index           =   1
         Left            =   135
         TabIndex        =   34
         Top             =   660
         Width           =   1365
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Cert. de Verificación"
         Height          =   195
         Index           =   2
         Left            =   120
         TabIndex        =   33
         Top             =   2355
         Width           =   1425
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Hoja de Verificación"
         Height          =   195
         Index           =   9
         Left            =   135
         TabIndex        =   32
         Top             =   1890
         Width           =   1425
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "R. Verificación"
         Height          =   195
         Index           =   25
         Left            =   5025
         TabIndex        =   15
         Top             =   2805
         Width           =   990
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "-"
         Height          =   195
         Index           =   29
         Left            =   7110
         TabIndex        =   14
         Top             =   2775
         Width           =   45
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Unidades"
         Height          =   195
         Index           =   30
         Left            =   510
         TabIndex        =   13
         Top             =   2805
         Width           =   675
      End
   End
   Begin VB.Image imagen 
      Height          =   480
      Left            =   9585
      Picture         =   "frmEquipoEdicionmantnimiento.frx":0CC0
      Top             =   0
      Width           =   480
   End
   Begin VB.Label lbltitulo 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00C0E0FF&
      Caption         =   "Verificación de Equipo"
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
      TabIndex        =   8
      Top             =   120
      Width           =   2325
   End
   Begin VB.Shape fondo 
      BackColor       =   &H00C0E0FF&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   510
      Left            =   0
      Top             =   0
      Width           =   10305
   End
End
Attribute VB_Name = "frmEquipoEdicionmantnimiento"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mvarlngPK As Long
Public booSilencioso As Boolean
Private mvarobjEquipo As clsEquipos
Private mvarenuTipoEdicion As enumTipoEdicion

Public Property Get TipoEdicion() As enumTipoEdicion

    TipoEdicion = mvarenuTipoEdicion

End Property

Public Property Let TipoEdicion(ByVal enuTipoEdicion As enumTipoEdicion)

    mvarenuTipoEdicion = enuTipoEdicion

End Property

Private Sub Form_Unload(Cancel As Integer)

    Set mvarobjEquipo = Nothing

End Sub


Public Property Get Equipo() As clsEquipos

    Set Equipo = mvarobjEquipo

End Property

Public Property Set Equipo(objEquipo As clsEquipos)

    Set mvarobjEquipo = objEquipo

End Property

Private Sub Command6_Click()

End Sub

Private Sub Form_Load_old()
    Dim titulo As String
    Dim oEQ_Cal As New clsEquipos_verificacion

    log (Me.Name)
    cargar_botones Me
    Call cargar_combos
    Call cabecera
    
    If oEQ_Cal.total_verificaciones(PK) = 1 Then ' si tiene verificaciones
        Call CARGAR ' se carga
        Call cargar_lista_verificaciones_hco
    Else
        ' se abre vacío
        Call marcoDatos_Verificacion_activo(False)
    End If
End Sub

Private Sub lblCampos_Click(Index As Integer)

End Sub

Private Sub txtdatos_GotFocus(Index As Integer)
    txtDatos(Index).BackColor = &H80C0FF
End Sub

Private Sub txtdatos_LostFocus(Index As Integer)
    txtDatos(Index).BackColor = vbWhite
End Sub

Private Sub cmbPeriVerificacion_Click(AREA As Integer)
' por ahora no se calcula en función de las fechas y periodicidad
'    If fecha_actual <> "" And cmbPeriVerificacion <> "" Then
'        fecha_proxima = calcular_fecha_proxima(fecha_actual.value, cmbPeriVerificacion.Text)
'    End If
End Sub

Private Sub fecha_actual_Change()
' por ahora no se calcula en función de las fechas y periodicidad
'    If fecha_actual <> "" And cmbPeriVerificacion <> "" Then
'        fecha_proxima = calcular_fecha_proxima(fecha_actual.value, cmbPeriVerificacion.Text)
'    End If
End Sub

Private Sub cmbTipoVerificacion_Change()
    If UCase(cmbTipoVerificacion.Text) = "EXTERNA" Then
        cmbCalibrador_interno.Limpiar
        cmbCalibrador_interno.Visible = False
        cmbOperador_Interno_Real.Limpiar
        cmbOperador_Interno_Real.Visible = False
        cmbCalibrador.Limpiar
        cmbCalibrador.cargar_datos
        cmbCalibrador.Visible = True
        cmbCalibrador.activar
        cmbOperador_Externo_Real.Limpiar
        cmbOperador_Externo_Real.cargar_datos
        cmbOperador_Externo_Real.Visible = True
        cmbOperador_Externo_Real.activar
    ElseIf UCase(cmbTipoVerificacion.Text) = "INTERNA" Then
        cmbCalibrador.Limpiar
        cmbCalibrador.Visible = False
        cmbOperador_Externo_Real.Limpiar
        cmbOperador_Externo_Real.Visible = False
        cmbCalibrador_interno.Limpiar
        cmbCalibrador_interno.cargar_datos
        cmbCalibrador_interno.Visible = True
        cmbCalibrador_interno.activar
        cmbOperador_Interno_Real.Limpiar
        cmbOperador_Interno_Real.cargar_datos
        cmbOperador_Interno_Real.Visible = True
        cmbOperador_Interno_Real.activar
    End If
End Sub

' botón que abre un cuadro de diálogo para seleccionar la plantilla excel de la verificación
Private Sub cmdExplorarDocumento_Click()
    On Error Resume Next
    cd.DialogTitle = "Abrir plantilla Excel"
    cd.ShowOpen
    If cd.FileName <> "" Then
        txtDatos(4) = cd.FileName ' Campo oculto para guardar ruta en la BD
        txtDatos(3) = cd.FileTitle ' Campo visible para mostrar en el formulario
    End If
End Sub

' botón que borra el documento de verificación
Private Sub cmdEliminarDocumento_Click()
    txtDatos(3) = ""
    txtDatos(4) = ""
End Sub

' botón que permite visualizar el archivo seleccionado
Private Sub cmdAbrirDocumento_Click()
    Call abrir_documento_excel(txtDatos(4))
End Sub

' al hacer doble click sobre un elemento se mostrará su documento excel
Private Sub lstLista_DblClick()
    'E0504-I
    If lstLista.ListItems.Count <> 0 Then
        Call abrir_documento_excel(lstLista.SelectedItem.SubItems(2))
    End If
    'E0504-F
End Sub

' botón que guarda los datos de una verificación en el histórico
Private Sub cmdGuardar_Verificacion_hco_Click()
    If datos_historico_correctos() Then
        If MsgBox("Va a introducir una nueva verificación. ¿Está seguro?", vbQuestion + vbYesNo, App.Title) = vbYes Then
            Dim lngEQC_HCO As Long
            Dim oEQC_HCO As New clsEquipos_Verificacion_Historico
            Dim ruta_archivo As String
            
            With oEQC_HCO
                .setEQUIPO_ID = PK
                .setFECHA = Format(fecha_actual, "yyyy-mm-dd hh:nn:ss")
'                .setFECHA = Format(Now, "yyyy-mm-dd hh:nn:ss")
                .setMODALIDAD = cmbTipoVerificacion.Text
                .setPROCEDIMIENTO = cmbProcedimientoVerificacion.Text
                If UCase(cmbTipoVerificacion.Text) = "INTERNA" Then
                    .setOPERADOR = cmbOperador_Interno_Real.getTEXTO
                    .setOPERADOR_ID = cmbOperador_Interno_Real.getPK_SALIDA
                ElseIf UCase(cmbTipoVerificacion.Text) = "EXTERNA" Then
                    .setOPERADOR = cmbOperador_Externo_Real.getTEXTO
                End If
                .setLIMITACIONES_USO = txtDatos(5)
                .setRANGO_MIN = txtDatos(1)
                .setRANGO_MAX = txtDatos(2)
                .setUNIDADES = cmbUnidad.getTEXTO
                .setCONFORME = chkConforme.value
                
                ruta_archivo = copiar_plantilla ' se copia la plantilla en otro documento para el histórico
                .setRUTA_DOCUMENTO = Replace(ruta_archivo, "\", "/")
            End With
            lngEQC_HCO = oEQC_HCO.Insertar
            
            Call cargar_lista_verificaciones_hco
            Call borrar_datos_verificacion
            
            fecha_actual = oEQC_HCO.getFECHA
            
            fecha_proxima = Equipos.calcular_fecha_proxima(fecha_actual.value, cmbPeriVerificacion.Text)
            booSilencioso = True
            Call cmdok_Click ' se guardan las fechas actual y próxima, y el resto de posibles cambios de la verificación
            booSilencioso = False
            
            If MsgBox("La verificación del equipo se insertó correctamente." & vbCrLf & _
                       "¿Desea imprimir la etiqueta de verificación?", vbYesNo + vbInformation, App.Title) = vbYes Then
                
                'strFirmaResponsable_verificacion = obtener_firma_responsable_verificacion
                Call imprimir_etiqueta(lstLista.SelectedItem.SubItems(1), lstLista.SelectedItem.SubItems(11))
                
            End If
            Call abrir_documento_excel(ruta_archivo)
            
            Set oEQC_HCO = Nothing
        Else
            Exit Sub
        End If
    End If
End Sub

' botón que permite eliminar una verificación del histórico
Private Sub cmdEliminarVerificacion_Historico_Click()
    If Not (lstLista.SelectedItem Is Nothing) Then
        If MsgBox("Se va a eliminar la verificación realizada el " & lstLista.SelectedItem.SubItems(3) & vbCrLf & _
                  "realizada por " & lstLista.SelectedItem.SubItems(5) & vbCrLf & _
                  "¿Está seguro?", vbYesNo + vbInformation, App.Title) = vbYes Then
            
            If Dir(lstLista.SelectedItem.SubItems(2), vbArchive) <> "" Then ' si el archivo existe
                Kill lstLista.SelectedItem.SubItems(2) ' se borra
            End If
            
            Dim oECH As New clsEquipos_Verificacion_Historico
            If oECH.Eliminar(lstLista.SelectedItem, lstLista.SelectedItem.SubItems(1)) Then
                Call cargar_lista_verificaciones_hco
                MsgBox "La verificación se ha eliminado correctamente.", vbOKOnly + vbInformation, App.Title
            End If
            Set oECH = Nothing
        End If
    Else
        MsgBox "Debe seleccionar la verificación que desea eliminar.", vbOKOnly + vbInformation, App.Title
    End If
End Sub

Private Sub cmdModificarVerificacion_Historico_Click()
    If Not (lstLista.SelectedItem Is Nothing) Then
        If datos_historico_correctos() Then
            If MsgBox("Se va a modificar la verificación realizada el " & lstLista.SelectedItem.SubItems(3) & vbCrLf & _
                      "realizada por " & lstLista.SelectedItem.SubItems(5) & vbCrLf & _
                      "¿Está seguro?", vbYesNo + vbInformation, App.Title) = vbYes Then
                Dim oECH As New clsEquipos_Verificacion_Historico
                With oECH
    '                .setEQUIPO_ID = PK
    '                .setFECHA = Format(Now, "yyyy-mm-dd hh:nn:ss")
                    .setMODALIDAD = cmbTipoVerificacion.Text
                    .setPROCEDIMIENTO = cmbProcedimientoVerificacion.Text
                    If UCase(cmbTipoVerificacion.Text) = "INTERNA" Then
                        .setOPERADOR = cmbOperador_Interno_Real.getTEXTO
                        .setOPERADOR_ID = cmbOperador_Interno_Real.getPK_SALIDA
                    ElseIf UCase(cmbTipoVerificacion.Text) = "EXTERNA" Then
                        .setOPERADOR = cmbOperador_Externo_Real.getTEXTO
                        .setOPERADOR_ID = cmbOperador_Externo_Real.getPK_SALIDA
                    End If
                    .setLIMITACIONES_USO = txtDatos(5)
                    .setRANGO_MIN = txtDatos(1)
                    .setRANGO_MAX = txtDatos(2)
                    .setUNIDADES = cmbUnidad.getTEXTO
                    .setCONFORME = chkConforme.value
                    .setRUTA_DOCUMENTO = Replace(lstLista.SelectedItem.SubItems(2), "\", "/")
                End With
                If oECH.Modificar(lstLista.SelectedItem, lstLista.SelectedItem.SubItems(1)) Then
                    Call cargar_lista_verificaciones_hco
                    Call borrar_datos_verificacion
                    MsgBox "La verificación se ha modificado en el histórico correctamente.", vbOKOnly + vbInformation, App.Title
                End If
                Set oECH = Nothing
            End If
        End If
    Else
        MsgBox "Debe seleccionar del histórico la verificación que desea modificar.", vbOKOnly + vbInformation, App.Title
    End If
End Sub

' botón que permite imprimir la etiqueta de verificación
Private Sub cmdEtiqueta_Click()

Dim objFrm As New frmEquipoEtiquetaVerificacion

    
    Set objFrm.Equipo = mvarobjEquipo
    objFrm.Show vbModal
    
    Unload objFrm
    Set objFrm = Nothing


'    If Not (lstLista.SelectedItem Is Nothing) Then
'        If lstLista.SelectedItem.Index = 1 Then ' sólo si está seleccionada la verificación más actual
'            Call imprimir_etiqueta(lstLista.SelectedItem.SubItems(1), lstLista.SelectedItem.SubItems(11))
'        Else
'            MsgBox "Debe estar seleccionada la verificación más actual del equipo" & vbCrLf & _
'                   "para generar su etiqueta.", vbOKOnly + vbInformation, App.Title
'        End If
'    End If
End Sub

Private Sub cmdok_Click()
    If datos_verificacion_correctos() Then
        Dim oEC As New clsEquipos_verificacion
        Dim EQUIPO_CAL As Long
        
        With oEC
            .setEQUIPO_ID = PK
            .setMODALIDAD_ID = cmbTipoVerificacion.BoundText
            .setPERIODICIDAD_ID = cmbPeriVerificacion.BoundText
            .setPROCEDIMIENTO_ID = IIf(cmbProcedimientoVerificacion.BoundText = "", 0, cmbProcedimientoVerificacion.BoundText)
            If UCase(cmbTipoVerificacion.Text) = "INTERNA" Then
                .setCALIBRADOR_INTERNO_ID = cmbCalibrador_interno.getPK_SALIDA
                .setCALIBRADOR_EXTERNO_ID = 0
            ElseIf UCase(cmbTipoVerificacion.Text) = "EXTERNA" Then
                .setCALIBRADOR_INTERNO_ID = 0
                .setCALIBRADOR_EXTERNO_ID = cmbCalibrador.getPK_SALIDA
            Else
                .setCALIBRADOR_INTERNO_ID = 0
                .setCALIBRADOR_EXTERNO_ID = 0
            End If
            .setFECHA_ACTUAL = Format(fecha_actual, "yyyy-mm-dd")
            .setFECHA_PROXIMA = Format(fecha_proxima, "yyyy-mm-dd")
            .setRANGO_MIN = txtDatos(1)
            .setRANGO_MAX = txtDatos(2)
            .setUNIDADES_ID = cmbUnidad.getPK_SALIDA
            .setRUTA_PLANTILLA = Replace(txtDatos(4), "\", "/")
        End With
        
        If oEC.total_verificaciones(PK) = 0 Then ' si no tiene verificaciones
            oEC.setEFECTIVA = 0
            If Not booSilencioso Then
                If MsgBox("Va a introducir los datos de la verificación. ¿Está seguro?", vbQuestion + vbYesNo, App.Title) = vbYes Then
                    EQUIPO_CAL = oEC.Insertar   ' se inserta
                Else
                    Exit Sub
                End If
            Else
                EQUIPO_CAL = oEC.Insertar
            End If
        Else                                    ' si tiene verificaciones
            oEC.setEFECTIVA = 1
            If Not booSilencioso Then
                If MsgBox("Va a modificar los datos de la verificación. ¿Está seguro?", vbQuestion + vbYesNo, App.Title) = vbYes Then
                    oEC.Modificar (PK)          ' se modifica
                    EQUIPO_CAL = PK
                Else
                    Exit Sub
                End If
            Else
                oEC.Modificar (PK)
                EQUIPO_CAL = PK
            End If
        End If
        
        Call marcoDatos_Verificacion_activo(True)
        frmEquipos_Detalle.datos_verificacion (PK) ' para actualizar la ventana frmEquipos_Detalle
        If Not booSilencioso Then
            MsgBox "La verificación del equipo se ha actualizado correctamente.", vbOKOnly + vbInformation, App.Title
        End If
    End If
End Sub

Private Sub cmdcancel_Click()
    Unload Me
End Sub

' ----------------- Funciones auxiliares del formulario ----------------
Private Sub cargar_combos()
    Dim oDECO As New clsDecodificadora
    
    oDECO.Cargar_Combo cmbTipoVerificacion, decodificadora.EQ_TIPO_CALIBRACION
    oDECO.Cargar_Combo cmbPeriVerificacion, decodificadora.EQ_periodicidad
    llenar_combo cmbCalibrador, New clsProveedor, 0, frmProveedores, ""
    llenar_combo cmbCalibrador_interno, New clsUsuarios, 0, Me, ""
    llenar_combo cmbUnidad, New clsUnidades, 0, Me, ""
    llenar_combo cmbOperador_Interno_Real, New clsUsuarios, 0, Me, ""
    llenar_combo cmbOperador_Externo_Real, New clsProveedor, 0, frmProveedores, ""
    
    ' Documentos de verificación
    Dim oCA_Doc As New clsCa_documentos
    Set cmbProcedimientoVerificacion.RowSource = oCA_Doc.Listado_Combo_procedimientos_verificacion()
    cmbProcedimientoVerificacion.ListField = "nombre" 'campo que veo
    cmbProcedimientoVerificacion.DataField = "id" 'campo asociado
    cmbProcedimientoVerificacion.BoundColumn = "id_documento" 'lo que realmente envia
    Set oCA_Doc = Nothing
    
    Set oDECO = Nothing
End Sub

Private Sub cabecera()
    With lstLista.ColumnHeaders
        .Add , , "ID", 0, lvwColumnLeft
        .Add , , "Fecha / Hora", 0, lvwColumnLeft
        .Add , , "Ruta_Excel", 0, lvwColumnLeft
        .Add , , "Fecha", 1200, lvwColumnLeft
        .Add , , "Modalidad", 1000, lvwColumnLeft
        .Add , , "Realizado por", 2000, lvwColumnLeft
        .Add , , "Limitaciones uso", 1500, lvwColumnLeft
        .Add , , "Rango Min", 1000, lvwColumnLeft
        .Add , , "Rango Máx", 1000, lvwColumnLeft
        .Add , , "Unidades", 1000, lvwColumnLeft
        .Add , , "Conforme", 1000, lvwColumnLeft
        .Add , , "ID_CALIBRADOR", 0, lvwColumnLeft
    End With
End Sub

' procedimiento que carga los datos de la verificación
Public Sub CARGAR()
    Dim oEquipo As New clsEquipos
    
    If oEquipo.Carga(PK) = True Then
        lbltitulo = "Verificación del Equipo : " & oEquipo.getNOMBRE
        Me.Caption = lbltitulo
        Dim oEC As New clsEquipos_verificacion
        If oEC.Carga(PK) Then
            With oEC
                cmbTipoVerificacion.BoundText = .getMODALIDAD_ID
                cmbCalibrador.MostrarElemento .getCALIBRADOR_EXTERNO_ID
                cmbCalibrador_interno.MostrarElemento .getCALIBRADOR_INTERNO_ID

                cmbOperador_Externo_Real.MostrarElemento .getCALIBRADOR_EXTERNO_ID
                cmbOperador_Interno_Real.MostrarElemento .getCALIBRADOR_INTERNO_ID

                cmbPeriVerificacion.BoundText = .getPERIODICIDAD_ID
                cmbProcedimientoVerificacion.BoundText = .getPROCEDIMIENTO_ID
                fecha_actual = .getFECHA_ACTUAL
                fecha_proxima = .getFECHA_PROXIMA
                txtDatos(1) = .getRANGO_MIN
                txtDatos(2) = .getRANGO_MAX
                cmbUnidad.MostrarElemento .getUNIDADES_ID
                If .getRUTA_PLANTILLA <> "" Then
                    txtDatos(3) = obtener_nombre_archivo(.getRUTA_PLANTILLA) ' sólo el nombre
                    txtDatos(4) = Replace(.getRUTA_PLANTILLA, "/", "\") ' toda la ruta
                End If
                
                Set oEC = Nothing
            End With
        End If
    End If
    Set oEquipo = Nothing
End Sub

' función que carga el histórico de verificaciónes
Private Sub cargar_lista_verificaciones_hco()
    Dim rs As ADODB.RecordSet
    Dim oECH As New clsEquipos_Verificacion_Historico
    
    lstLista.ListItems.Clear
    Set rs = oECH.Listado(PK)
    If rs.RecordCount <> 0 Then
        Do
            With lstLista.ListItems.Add(, , Format(rs(0), "0000"))
                .SubItems(1) = Format(rs(1), "yyyy-mm-dd hh:nn:ss")
                .SubItems(2) = Replace(rs(2), "/", "\")
                .SubItems(3) = Format(rs(3), "dd/mm/yyyy")
                .SubItems(4) = rs(4)
                .SubItems(5) = rs(5)
                .SubItems(6) = rs(6)
                .SubItems(7) = rs(7)
                .SubItems(8) = rs(8)
                .SubItems(9) = rs(9)
                .SubItems(10) = rs(10)
                .SubItems(11) = rs(11)
            End With
            rs.MoveNext
        Loop Until rs.EOF
    End If
    Set oECH = Nothing
End Sub

' función que comprueba que los datos de la verificación son correctos
Private Function datos_verificacion_correctos() As Boolean
'    If Trim(txtDatos(3)) = "" Then ' excel de verificación
'        MsgBox "Debe indicar una plantilla de verificacion.", vbInformation, App.Title
'        txtDatos(3).SetFocus
'        datos_verificacion_correctos = False
'        Exit Function
'    End If
    datos_verificacion_correctos = True
End Function

' función que comprueba que los datos de la verificación histórico son correctos
Private Function datos_historico_correctos() As Boolean
    If UCase(cmbTipoVerificacion.Text) = "INTERNA" Then
        If Trim(cmbOperador_Interno_Real.getTEXTO) = "" Then
            MsgBox "Debe indicar el operador de la verificación.", vbInformation, App.Title
            cmbOperador_Interno_Real.SetFocus
            datos_historico_correctos = False
            Exit Function
        End If
    ElseIf UCase(cmbTipoVerificacion.Text) = "EXTERNA" Then
        If Trim(cmbOperador_Externo_Real.getTEXTO) = "" Then
            MsgBox "Debe indicar el operador de la verificación.", vbInformation, App.Title
            cmbOperador_Externo_Real.SetFocus
            datos_historico_correctos = False
            Exit Function
        End If
    End If
    
    If Trim(txtDatos(5)) = "" Then ' limitaciones de uso
        MsgBox "Debe indicar limitaciones de uso de la verificación.", vbInformation, App.Title
        txtDatos(5).SetFocus
        datos_historico_correctos = False
        Exit Function
    End If

    If Trim(txtDatos(3)) = "" Then ' documento excel
        MsgBox "Debe indicar documento excel de plantilla.", vbInformation, App.Title
        txtDatos(3).SetFocus
        datos_historico_correctos = False
        Exit Function
    End If
    
    datos_historico_correctos = True
End Function

' procedimiento que restablece los campos de la verificación
Private Sub borrar_datos_verificacion()
    cmbOperador_Interno_Real.Limpiar
    cmbOperador_Externo_Real.Limpiar
    txtDatos(5).Text = ""
    chkConforme.value = 0
End Sub

Private Sub marcoDatos_Verificacion_activo(booEstado As Boolean)
    marcoDatos_Verificacion.Enabled = booEstado
    If booEstado Then
        cmbOperador_Externo_Real.activar
        cmbOperador_Interno_Real.activar
    Else
        cmbOperador_Externo_Real.desactivar
        cmbOperador_Interno_Real.desactivar
    End If
    chkConforme.Enabled = booEstado
    cmdGuardar_Verificacion_hco.Enabled = booEstado
End Sub

' función que permite abrir el documento excel pasado por parámetro
Private Function abrir_documento_excel(strRuta As String) As Boolean
    Dim destino As String
    Dim r As Long
    
On Error GoTo fallo
    
    If Len(Trim(strRuta)) > 0 Then
        destino = Replace(strRuta, "/", "\")
        If destino = "" Then
            Exit Function
        End If
        If Dir(destino) <> "" Then
            r = Shell("rundll32.exe url.dll,FileProtocolHandler " & destino, vbMaximizedFocus)
        Else
            MsgBox "El documento se ha eliminado o movido de la ruta almacenada:" & vbCrLf & _
                   destino, vbCritical, App.Title
        End If
    Else
        MsgBox "No hay ningún documento para mostrar.", vbCritical, App.Title
    End If

    Exit Function
    
fallo:
    MsgBox "Error al abrir el documento.", vbCritical, App.Title
End Function

' Botón que copia la plantilla excel de verificación en un excel para la introducción de datos concretos
Private Function copiar_plantilla() As String
    On Error GoTo fallo
    Dim doc As String, strNombreArchivo As String
    
    strNombreArchivo = Format(PK, "0000") & "_" & Format(Now, "yyyymmdd_hhnnss") & ".xls"
    If UCase(USUARIO.getNOMBRE) = "PRUEBA" Then
        doc = ReadINI(App.Path + "\config.ini", "documentos", "ruta") & "\prueba\"
    Else
        doc = ReadINI(App.Path + "\config.ini", "documentos", "ruta")
    End If
    doc = doc & "\Equipos\Verificacion\" & Year(Date)
    
    If Dir(doc, vbDirectory) <> "" Then
        doc = doc & "\" & strNombreArchivo
        FileCopy txtDatos(4), doc
        copiar_plantilla = doc
    Else
        MkDir doc
    End If
    
    Me.MousePointer = 0
    
    Exit Function
    
fallo:
    Me.MousePointer = 0
    MsgBox "Error al generar el documento.", vbCritical, App.Title

End Function

' función que devuelve el nombre de archivo contenido en la ruta pasada por parámetro
Private Function obtener_nombre_archivo(strRuta As String) As String
    Dim arrRuta() As String
    
    arrRuta = Split(strRuta, "/")
    obtener_nombre_archivo = arrRuta(UBound(arrRuta))
    
End Function

Private Sub imprimir_etiqueta(strFecha_Verificacion As String, lngOperador_ID As Long)
    Dim prnPrinter As Printer
    
On Error GoTo trataError

    ' se mira si el equipo tiene impresora de etiquetas
    Dim oParametro As New clsParametros
    If Not oParametro.Carga(parametros.IMPRESORA_ETIQUETAS, USUARIO.getUSO) Then
        MsgBox "Este equipo no tiene asignada impresora de etiquetas.", vbCritical, App.Title
        Exit Sub
    End If
    log ("Comienzo impresion de etiqueta de verificación de equipo")
    For Each prnPrinter In Printers
        If prnPrinter.DeviceName = oParametro.getVALOR Then
            Set Printer = prnPrinter
            Exit For
        End If
    Next
    
    'Call Firmas.copiar_firma_responsable_verificacion(picture1, lngOperador_ID)
    
    With frmReport
        .iniciar
        .informe = "rptEquipos_ETIQUETA_Verificacion"
        .criterio = "{equipos.ID_EQUIPO} = " & CLng(PK) & _
                    "and {equipos_verificacion_historico.FECHA}= datetime('" & strFecha_Verificacion & "')"
        .imprimir = True
        .generar
        .Visible = False
    End With
    log ("Final impresion de etiqueta de verificación de equipo")
    
    Exit Sub
    
trataError:
    MsgBox "Error al imprimir la etiqueta de verificación.", vbCritical, Err.Description
End Sub

Public Property Get PK() As Long

    PK = mvarlngPK

End Property

Public Property Let PK(ByVal lngPK As Long)

    mvarlngPK = lngPK

End Property
