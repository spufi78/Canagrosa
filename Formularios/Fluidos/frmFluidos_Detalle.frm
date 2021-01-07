VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{CEA3B7F6-2847-4E5E-A551-DB7A62489D44}#46.0#0"; "miCombo.ocx"
Begin VB.Form frmFluidos_Detalle 
   BackColor       =   &H00C0C0C0&
   Caption         =   "Mantenimiento de Fluidos"
   ClientHeight    =   8880
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10815
   Icon            =   "frmFluidos_Detalle.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   8880
   ScaleWidth      =   10815
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame3 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   "Clasificación Fluido AIM (Aplicación control de procesos)"
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
      Height          =   2280
      Left            =   45
      TabIndex        =   31
      Top             =   5625
      Width           =   10725
      Begin pryCombo.miCombo cmbCentro 
         Height          =   375
         Left            =   1305
         TabIndex        =   33
         Top             =   630
         Width           =   9345
         _ExtentX        =   16484
         _ExtentY        =   661
      End
      Begin pryCombo.miCombo cmbTipoEnsayo 
         Height          =   375
         Left            =   1305
         TabIndex        =   34
         Top             =   1035
         Width           =   9345
         _ExtentX        =   16484
         _ExtentY        =   661
      End
      Begin pryCombo.miCombo cmbSeccion 
         Height          =   375
         Left            =   1305
         TabIndex        =   38
         Top             =   1440
         Width           =   9345
         _ExtentX        =   16484
         _ExtentY        =   661
      End
      Begin pryCombo.miCombo cmbEstacion 
         Height          =   375
         Left            =   1305
         TabIndex        =   40
         Top             =   1845
         Width           =   9345
         _ExtentX        =   16484
         _ExtentY        =   661
      End
      Begin pryCombo.miCombo cmbPrograma 
         Height          =   375
         Left            =   1305
         TabIndex        =   32
         Top             =   240
         Width           =   9345
         _ExtentX        =   16484
         _ExtentY        =   661
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Estación"
         Height          =   195
         Index           =   9
         Left            =   90
         TabIndex        =   41
         Top             =   1890
         Width           =   615
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Sección"
         Height          =   195
         Index           =   7
         Left            =   90
         TabIndex        =   39
         Top             =   1485
         Width           =   585
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Programa"
         Height          =   195
         Index           =   20
         Left            =   90
         TabIndex        =   37
         Top             =   270
         Width           =   675
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Centro"
         Height          =   195
         Index           =   18
         Left            =   90
         TabIndex        =   36
         Top             =   675
         Width           =   465
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Tipo de Ensayo"
         Height          =   195
         Index           =   19
         Left            =   90
         TabIndex        =   35
         Top             =   1080
         Width           =   1110
      End
   End
   Begin VB.Frame Frame9 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   "Informe"
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
      Height          =   720
      Left            =   30
      TabIndex        =   28
      Top             =   4875
      Width           =   10755
      Begin VB.CheckBox chkGrado 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Incluir Columna de Grado"
         Height          =   285
         Left            =   8145
         TabIndex        =   30
         Top             =   255
         Width           =   2475
      End
      Begin VB.CheckBox chkVB 
         BackColor       =   &H00C0C0C0&
         Caption         =   "No Incluir VºBº"
         Height          =   285
         Left            =   4770
         TabIndex        =   11
         Top             =   255
         Width           =   2475
      End
      Begin VB.CheckBox chkCP 
         BackColor       =   &H00C0C0C0&
         Caption         =   "No Incluir detalle de contaminación por partículas"
         Height          =   285
         Left            =   180
         TabIndex        =   10
         Top             =   255
         Width           =   4185
      End
   End
   Begin VB.CommandButton cmdHistorialCambios 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Historial Cambios"
      Height          =   870
      Left            =   7155
      Style           =   1  'Graphical
      TabIndex        =   27
      Top             =   7965
      Width           =   1365
   End
   Begin VB.CommandButton cmdInsertaNorma 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Buscar Especificación de Control"
      Height          =   870
      Left            =   90
      Picture         =   "frmFluidos_Detalle.frx":1272
      Style           =   1  'Graphical
      TabIndex        =   26
      Top             =   7965
      Visible         =   0   'False
      Width           =   2895
   End
   Begin VB.Frame Frame2 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   "Baño Asociado"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   645
      Left            =   45
      TabIndex        =   19
      Top             =   4200
      Width           =   10755
      Begin pryCombo.miCombo cmbbano 
         Height          =   345
         Left            =   1755
         TabIndex        =   9
         Top             =   225
         Width           =   8970
         _ExtentX        =   15822
         _ExtentY        =   609
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Baño"
         Height          =   195
         Index           =   8
         Left            =   90
         TabIndex        =   20
         Top             =   270
         Width           =   375
      End
   End
   Begin VB.CommandButton cmdSalir 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Salir"
      Height          =   870
      Left            =   9720
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   7965
      Width           =   1050
   End
   Begin VB.CommandButton cmdok 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Aceptar"
      Height          =   870
      Left            =   8595
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   7965
      Width           =   1050
   End
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   "Datos del Fluido"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   3390
      Left            =   45
      TabIndex        =   14
      Top             =   765
      Width           =   10755
      Begin VB.TextBox txtDatos 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Index           =   2
         Left            =   1755
         TabIndex        =   8
         Top             =   3030
         Width           =   7950
      End
      Begin VB.TextBox txtDatos 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Index           =   1
         Left            =   1755
         TabIndex        =   7
         Top             =   2700
         Width           =   7950
      End
      Begin VB.TextBox txtDatos 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Index           =   0
         Left            =   1755
         TabIndex        =   0
         Top             =   270
         Width           =   7950
      End
      Begin pryCombo.miCombo cmbsub 
         Height          =   345
         Left            =   1755
         TabIndex        =   2
         Top             =   945
         Width           =   8970
         _ExtentX        =   15822
         _ExtentY        =   609
      End
      Begin pryCombo.miCombo cmbnorma 
         Height          =   345
         Left            =   1755
         TabIndex        =   5
         Top             =   1995
         Width           =   8970
         _ExtentX        =   15822
         _ExtentY        =   609
      End
      Begin MSDataListLib.DataCombo cmbcla 
         Height          =   315
         Left            =   1755
         TabIndex        =   1
         Top             =   585
         Width           =   7935
         _ExtentX        =   13996
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         Style           =   2
         Text            =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin pryCombo.miCombo cmbcontrol 
         Height          =   345
         Left            =   1755
         TabIndex        =   3
         Top             =   1305
         Width           =   8970
         _ExtentX        =   15822
         _ExtentY        =   609
      End
      Begin pryCombo.miCombo cmbcontrol2 
         Height          =   345
         Left            =   1755
         TabIndex        =   4
         Top             =   1650
         Width           =   8970
         _ExtentX        =   15822
         _ExtentY        =   609
      End
      Begin pryCombo.miCombo cmbPNT 
         Height          =   345
         Left            =   1755
         TabIndex        =   6
         Top             =   2340
         Width           =   8970
         _ExtentX        =   15822
         _ExtentY        =   609
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "P.N.T."
         Height          =   195
         Index           =   10
         Left            =   90
         TabIndex        =   42
         Top             =   2385
         Width           =   465
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Esp. de Control (2)"
         Height          =   195
         Index           =   6
         Left            =   90
         TabIndex        =   29
         Top             =   1695
         Width           =   1305
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Esp. de Control (1)"
         Height          =   195
         Index           =   5
         Left            =   90
         TabIndex        =   25
         Top             =   1350
         Width           =   1305
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Normativa Referencia"
         Height          =   195
         Index           =   1
         Left            =   90
         TabIndex        =   24
         Top             =   3030
         Width           =   1545
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Normativa Aplicable"
         Height          =   195
         Index           =   0
         Left            =   90
         TabIndex        =   23
         Top             =   2715
         Width           =   1410
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Norma"
         Height          =   195
         Index           =   11
         Left            =   90
         TabIndex        =   18
         Top             =   2040
         Width           =   465
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Subclasificación"
         Height          =   195
         Index           =   4
         Left            =   90
         TabIndex        =   17
         Top             =   990
         Width           =   1155
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Descripción"
         Height          =   195
         Index           =   2
         Left            =   90
         TabIndex        =   15
         Top             =   315
         Width           =   840
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Clasificación"
         Height          =   195
         Index           =   3
         Left            =   90
         TabIndex        =   16
         Top             =   630
         Width           =   885
      End
   End
   Begin VB.Label lblsubtitulo 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Caption         =   $"frmFluidos_Detalle.frx":1B3C
      ForeColor       =   &H80000008&
      Height          =   240
      Left            =   90
      TabIndex        =   22
      Top             =   375
      Width           =   9645
   End
   Begin VB.Image imagen 
      Height          =   480
      Left            =   10215
      Picture         =   "frmFluidos_Detalle.frx":1BCA
      Top             =   135
      Width           =   480
   End
   Begin VB.Label lbltitulo 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "Ficha de Fluido"
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
      TabIndex        =   21
      Top             =   75
      Width           =   1620
   End
   Begin VB.Shape fondo 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   735
      Left            =   0
      Top             =   0
      Width           =   11025
   End
End
Attribute VB_Name = "frmFluidos_Detalle"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public PK As Long
Private Sub cmdHistorialCambios_Click()
    If PK <> 0 Then
        frmHistorialCambios.PK_TIPO = HC_TIPOS.HC_FLUIDO
        frmHistorialCambios.PK_ID = PK
        frmHistorialCambios.PK_TITULO = "Fluido " & txtdatos(0)
        frmHistorialCambios.Show 1
    End If
End Sub

Private Sub cmdok_Click()
    On Error GoTo fallo
    If validar = True Then
      Dim oFluido As New clsFluidos_ficha
      With oFluido
        .setDESCRIPCION = txtdatos(0)
        .setBANO_ID = cmbBano.getPK_SALIDA
        .setNORMA_ID = cmbnorma.getPK_SALIDA
        .setNORMA_CONTROL_ID = cmbcontrol.getPK_SALIDA
        .setNORMA_CONTROL_ID2 = cmbcontrol2.getPK_SALIDA
        .setDOCUMENTO_ID = cmbPNT.getPK_SALIDA
        .setSUBCLASIFICACION_ID = cmbsub.getPK_SALIDA
        .setTIPO_MUESTRA_ID = cmbcla.BoundText
        .setNORMATIVA_APLICABLE = txtdatos(1)
        .setNORMATIVA_REFERENCIA = txtdatos(2)
        
        .setINFORME_CONTAMINACION = chkCP.Value
        .setINFORME_VB = chkVB.Value
        .setINFORME_GRADO = chkGrado.Value
        
        .setAIM_PROGRAMA_ID = cmbPrograma.getPK_SALIDA
        .setAIM_CENTRO_ID = cmbCentro.getPK_SALIDA
        .setAIM_TIPO_ENSAYO_ID = cmbTipoEnsayo.getPK_SALIDA
        .setAIM_SECCION_ID = cmbSeccion.getPK_SALIDA
        .setAIM_ESTACION_ID = cmbEstacion.getPK_SALIDA
      End With
      Dim ohc As New clsHistorial_cambios
      If PK = 0 Then
        If MsgBox("Va a introducir el nuevo fluido. ¿Esta seguro?", vbQuestion + vbYesNo, App.Title) = vbYes Then
            PK = oFluido.Insertar
            If PK <> 0 Then
                With ohc
                    .setTIPO = HC_TIPOS.HC_FLUIDO
                    .setIDENTIFICADOR = PK
                    .setIDENTIFICADOR_TEXTO = txtdatos(0)
                    .setUSUARIO_ID = USUARIO.getID_EMPLEADO
                    .setMOTIVO = HC_CREACION
                    .Insertar
                End With
            End If
        Else
            Exit Sub
        End If
      Else
        If MsgBox("Va a modificar el fluido. ¿Esta seguro?", vbQuestion + vbYesNo, App.Title) = vbYes Then
            frmMotivo.lbltitulo = "Indique detalladamente el motivo de modificación del fluido."
            frmMotivo.Show 1
            If Trim(MOTIVO) = "" Then
                MsgBox "Para modificar los datos es necesario introducir el motivo de la modificación.", vbInformation, App.Title
                Exit Sub
            End If
            If oFluido.Modificar(PK) = False Then
                Exit Sub
            Else
                If PK <> 0 Then
                    With ohc
                        .setTIPO = HC_TIPOS.HC_FLUIDO
                        .setIDENTIFICADOR = PK
                        .setIDENTIFICADOR_TEXTO = txtdatos(0)
                        .setUSUARIO_ID = USUARIO.getID_EMPLEADO
                        .setMOTIVO = Trim(MOTIVO)
                        .Insertar
                    End With
                End If
            End If
        Else
            Exit Sub
        End If
      End If
      Set ohc = Nothing
      MsgBox "Actualizaciones realizadas correctamente.", vbOKOnly + vbInformation, App.Title
      Unload Me
    End If
    Exit Sub
fallo:
    error_grave ("Error al insertar el fluido : " & Err.Description)
End Sub


Private Sub cmdInsertaNorma_Click()
    gID = 0
    frmCA_Listado_Normas.VINCULAR = True
    frmCA_Listado_Normas.Show 1
    If gID <> 0 Then
        cmbcontrol.MostrarElemento gID
'        Dim oNorma As New clsCa_normas
'        If oNorma.Carga(CLng(gID)) Then
'            With listaDocumentacion.ListItems.Add(, , gID)
'                .SubItems(1) = oNorma.getNOMBRE
'                .SubItems(2) = "NORMA"
'            End With
'        End If
    End If
End Sub


Private Sub Form_Load()
    log (Me.Name)
    cargar_botones Me
    cargar_combos
    
    If PK <> 0 Then
        cargar_datos
    End If
End Sub
Private Sub cmdSalir_Click()
    Unload Me
End Sub
Private Sub Form_Unload(Cancel As Integer)
    PK = 0
End Sub

Private Sub txtdatos_GotFocus(Index As Integer)
    txtdatos(Index).BackColor = &H80C0FF
    txtdatos(Index).SelStart = 0
    txtdatos(Index).SelLength = Len(txtdatos(Index))
End Sub
Private Sub txtDatos_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys "{Tab}", True
        KeyAscii = 0 ' Para evitar el "bip" del sistema
    End If
End Sub
Private Sub txtdatos_LostFocus(Index As Integer)
    txtdatos(Index).BackColor = vbWhite
End Sub
Public Function validar() As Boolean
    validar = True
    If Trim(txtdatos(0)) = "" Then
        MsgBox "Debe darle un nombre al fluido.", vbExclamation, App.Title
        validar = False
        Exit Function
    End If
    If cmbBano.getPK_SALIDA = 0 Then
        MsgBox "Debe asignar un baño al fluido.", vbExclamation, App.Title
        validar = False
        Exit Function
    End If
    If cmbnorma.getTEXTO = "" Then
        MsgBox "Debe asignar una norma al fluido.", vbExclamation, App.Title
        validar = False
        Exit Function
    End If
    If cmbcla.Text = "" Then
        MsgBox "Debe asignar una clasificación al fluido.", vbExclamation, App.Title
        validar = False
        Exit Function
    End If
End Function

Public Sub cargar_datos()
    Dim oFluido As New clsFluidos_ficha
    With oFluido
        If .Carga(PK) = True Then
            lbltitulo = "Modificación Fluido : " & .getDESCRIPCION
            Me.Caption = lbltitulo
            txtdatos(0) = .getDESCRIPCION
            cmbcla.BoundText = .getTIPO_MUESTRA_ID
            cmbsub.MostrarElemento .getSUBCLASIFICACION_ID
            cmbnorma.MostrarElemento .getNORMA_ID
            cmbBano.MostrarElemento .getBANO_ID
            cmbcontrol.MostrarElemento .getNORMA_CONTROL_ID
            cmbcontrol2.MostrarElemento .getNORMA_CONTROL_ID2
            cmbPNT.MostrarElemento .getDOCUMENTO_ID
            txtdatos(1) = .getNORMATIVA_APLICABLE
            txtdatos(2) = .getNORMATIVA_REFERENCIA
            chkCP.Value = .getINFORME_CONTAMINACION
            chkVB.Value = .getINFORME_VB
            chkGrado.Value = .getINFORME_GRADO
            
            cmbPrograma.MostrarElemento .getAIM_PROGRAMA_ID
            cmbCentro.MostrarElemento .getAIM_CENTRO_ID
            cmbTipoEnsayo.MostrarElemento .getAIM_TIPO_ENSAYO_ID
            cmbSeccion.MostrarElemento .getAIM_SECCION_ID
            cmbEstacion.MostrarElemento .getAIM_ESTACION_ID
        End If
    End With
    Set oFluido = Nothing
End Sub
Public Sub cargar_combos()
    Dim otm As New clsTipos_muestra
    Set cmbcla.RowSource = otm.Listado_Fluidos
    cmbcla.ListField = "nombre"
    cmbcla.BoundColumn = "id_tipo_muestra"
    llenar_combo cmbcontrol, New clsCa_normas, 0, frmCA_Normas, ""
    llenar_combo cmbcontrol2, New clsCa_normas, 0, frmCA_Normas, ""
    llenar_combo cmbsub, New clsFluidos_subclasificacion, 0, frmClientes, ""
    llenar_combo cmbnorma, New clsFluidos_normas, 0, frmClientes, ""
    llenar_combo cmbBano, New clsBanos, 0, frmBANO_Detalle, ""
    llenar_combo cmbPNT, New clsCa_documentos, 0, frmCA_Documento, " ANULADO = 0 "
    ' Nuevos campos AIM
    Dim oDeco As New clsDecodificadora
    oDeco.cargar_mi_combo cmbPrograma, DECODIFICADORA.FLUIDOS_PROGRAMAS
    oDeco.cargar_mi_combo cmbCentro, DECODIFICADORA.FLUIDOS_CENTROS
    oDeco.cargar_mi_combo cmbTipoEnsayo, DECODIFICADORA.FLUIDOS_TIPOS_ENSAYOS
    oDeco.cargar_mi_combo cmbSeccion, DECODIFICADORA.FLUIDOS_SECCIONES
    oDeco.cargar_mi_combo cmbEstacion, DECODIFICADORA.FLUIDOS_ESTACIONES
    Set oDeco = Nothing
    
End Sub
