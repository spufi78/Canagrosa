VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{CEA3B7F6-2847-4E5E-A551-DB7A62489D44}#35.0#0"; "miCombo.ocx"
Begin VB.Form frmEquipos_Detalle_Verificacion 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   7845
   ClientLeft      =   2415
   ClientTop       =   2475
   ClientWidth     =   10140
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "frmEquipos_Detalle_Verificacion.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7845
   ScaleWidth      =   10140
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdEtiqueta 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Etiqueta"
      Height          =   870
      Left            =   2700
      Style           =   1  'Graphical
      TabIndex        =   40
      Top             =   6975
      Width           =   1050
   End
   Begin VB.CommandButton cmdEliminarVerificacion 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Eliminar verificación"
      Height          =   870
      Left            =   90
      Style           =   1  'Graphical
      TabIndex        =   38
      Top             =   6975
      Width           =   1050
   End
   Begin VB.CommandButton cmdEliminarVerificacion_Historico 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Eliminar verificación de histórico"
      Height          =   870
      Left            =   1215
      Style           =   1  'Graphical
      TabIndex        =   37
      Top             =   6975
      Width           =   1050
   End
   Begin VB.Frame frmVerificacionesAsignadas 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   "Verificaciones asignadas"
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
      Height          =   1275
      Left            =   45
      TabIndex        =   33
      Top             =   495
      Width           =   10080
      Begin VB.CommandButton cmdModificarVerificacion 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Modificar"
         Height          =   330
         Left            =   8820
         Style           =   1  'Graphical
         TabIndex        =   36
         Top             =   630
         Width           =   1140
      End
      Begin VB.CommandButton cmdNuevaVerificacion 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Limpiar datos"
         Height          =   330
         Left            =   8820
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   225
         Width           =   1140
      End
      Begin MSComctlLib.ListView listaVerificacionesAsignadas 
         Height          =   885
         Left            =   135
         TabIndex        =   35
         Top             =   225
         Width           =   8475
         _ExtentX        =   14949
         _ExtentY        =   1561
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
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   0
      End
   End
   Begin VB.Frame frmDatosVerificacion 
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
      Height          =   1005
      Left            =   45
      TabIndex        =   22
      Top             =   4185
      Width           =   10080
      Begin VB.CheckBox chkConforme 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Conforme"
         Height          =   240
         Left            =   7605
         TabIndex        =   11
         Top             =   270
         Width           =   1320
      End
      Begin VB.CommandButton cmdGuardar_Verificacion_hco 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Guardar"
         Height          =   330
         Left            =   9000
         Style           =   1  'Graphical
         TabIndex        =   13
         Top             =   225
         Width           =   960
      End
      Begin VB.TextBox txtDatos 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   330
         Index           =   4
         Left            =   1620
         MaxLength       =   100
         TabIndex        =   10
         Top             =   585
         Width           =   8340
      End
      Begin pryCombo.miCombo cmbOperador_Interno_Real 
         Height          =   330
         Left            =   1620
         TabIndex        =   9
         Top             =   225
         Width           =   5655
         _ExtentX        =   9975
         _ExtentY        =   582
      End
      Begin pryCombo.miCombo cmbOperador_Externo_Real 
         Height          =   330
         Left            =   1620
         TabIndex        =   23
         Top             =   225
         Width           =   5655
         _ExtentX        =   9975
         _ExtentY        =   582
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Limitaciones uso"
         Height          =   195
         Index           =   5
         Left            =   135
         TabIndex        =   25
         Top             =   630
         Width           =   1170
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Realizado por"
         Height          =   195
         Index           =   6
         Left            =   135
         TabIndex        =   24
         Top             =   270
         Width           =   975
      End
   End
   Begin VB.Frame frmHistoricoVerificaciones 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   "Historico de verificaciones"
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
      Height          =   1680
      Left            =   45
      TabIndex        =   20
      Top             =   5220
      Width           =   10080
      Begin MSComctlLib.ListView lstLista 
         Height          =   1335
         Left            =   135
         TabIndex        =   21
         Top             =   270
         Width           =   9825
         _ExtentX        =   17330
         _ExtentY        =   2355
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
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   0
      End
   End
   Begin VB.CommandButton cmdcancel 
      BackColor       =   &H00E0E0E0&
      Caption         =   "ESC-Salir"
      Height          =   870
      Left            =   9045
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   6975
      Width           =   1050
   End
   Begin VB.CommandButton cmdok 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Nueva"
      Height          =   870
      Left            =   7965
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   6975
      Width           =   1050
   End
   Begin VB.Frame frmVerificacion 
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
      Height          =   2355
      Left            =   45
      TabIndex        =   16
      Top             =   1800
      Width           =   10080
      Begin VB.TextBox txtDatos 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   330
         Index           =   0
         Left            =   9000
         MaxLength       =   50
         TabIndex        =   39
         Top             =   0
         Visible         =   0   'False
         Width           =   690
      End
      Begin VB.TextBox txtDatos 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   330
         Index           =   2
         Left            =   5580
         MaxLength       =   50
         TabIndex        =   7
         Top             =   1485
         Width           =   690
      End
      Begin VB.TextBox txtDatos 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   330
         Index           =   1
         Left            =   4770
         MaxLength       =   50
         TabIndex        =   6
         Top             =   1485
         Width           =   690
      End
      Begin MSComCtl2.DTPicker fecha_actual 
         Height          =   345
         Left            =   1680
         TabIndex        =   4
         Top             =   1485
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
         Format          =   57802753
         CurrentDate     =   2
         MinDate         =   2
      End
      Begin MSComCtl2.DTPicker fecha_proxima 
         Height          =   345
         Left            =   1680
         TabIndex        =   5
         Top             =   1890
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
         Format          =   57802753
         CurrentDate     =   2
         MinDate         =   2
      End
      Begin pryCombo.miCombo cmbVerificador_interno 
         Height          =   330
         Left            =   1665
         TabIndex        =   3
         Top             =   1125
         Width           =   5385
         _ExtentX        =   9499
         _ExtentY        =   582
      End
      Begin MSDataListLib.DataCombo cmbTipoVerificacion 
         Height          =   315
         Left            =   1680
         TabIndex        =   0
         Top             =   405
         Width           =   2925
         _ExtentX        =   5159
         _ExtentY        =   556
         _Version        =   393216
         Style           =   2
         Text            =   ""
      End
      Begin MSDataListLib.DataCombo cmbProcedimientoVerificacion 
         Height          =   315
         Left            =   1680
         TabIndex        =   2
         Top             =   765
         Width           =   8010
         _ExtentX        =   14129
         _ExtentY        =   556
         _Version        =   393216
         Style           =   2
         Text            =   ""
      End
      Begin MSDataListLib.DataCombo cmbPeriVerificacion 
         Height          =   315
         Left            =   6540
         TabIndex        =   1
         Top             =   405
         Width           =   3150
         _ExtentX        =   5556
         _ExtentY        =   556
         _Version        =   393216
         Style           =   2
         Text            =   ""
      End
      Begin pryCombo.miCombo cmbUnidad 
         Height          =   330
         Left            =   7155
         TabIndex        =   8
         Top             =   1485
         Width           =   2760
         _ExtentX        =   4868
         _ExtentY        =   582
      End
      Begin pryCombo.miCombo cmbVerificador 
         Height          =   330
         Left            =   1665
         TabIndex        =   34
         Top             =   1125
         Width           =   5385
         _ExtentX        =   9499
         _ExtentY        =   582
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Unidades"
         Height          =   195
         Index           =   30
         Left            =   6435
         TabIndex        =   32
         Top             =   1530
         Width           =   675
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "-"
         Height          =   195
         Index           =   29
         Left            =   5505
         TabIndex        =   31
         Top             =   1530
         Width           =   45
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "R. Verificación"
         Height          =   195
         Index           =   25
         Left            =   3690
         TabIndex        =   30
         Top             =   1530
         Width           =   1035
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Periodo Verificación"
         Height          =   195
         Index           =   1
         Left            =   5085
         TabIndex        =   29
         Top             =   450
         Width           =   1410
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Modalidad"
         Height          =   195
         Index           =   3
         Left            =   120
         TabIndex        =   28
         Top             =   450
         Width           =   735
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Procedimiento"
         Height          =   195
         Index           =   4
         Left            =   120
         TabIndex        =   27
         Top             =   810
         Width           =   1005
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Responsable"
         Height          =   195
         Index           =   8
         Left            =   120
         TabIndex        =   26
         Top             =   1170
         Width           =   930
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "F. Próx. Verificación"
         Height          =   195
         Index           =   0
         Left            =   120
         TabIndex        =   19
         Top             =   1935
         Width           =   1410
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "F. Actual Verificación"
         Height          =   195
         Index           =   10
         Left            =   120
         TabIndex        =   18
         Top             =   1575
         Width           =   1500
      End
   End
   Begin VB.Image imagen 
      Height          =   360
      Left            =   9630
      Picture         =   "frmEquipos_Detalle_Verificacion.frx":000C
      Top             =   0
      Width           =   360
   End
   Begin VB.Label lbltitulo 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00C0FFFF&
      Caption         =   "Equipo"
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
      TabIndex        =   17
      Top             =   120
      Width           =   750
   End
   Begin VB.Shape fondo 
      BackColor       =   &H00C0FFFF&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   510
      Left            =   0
      Top             =   0
      Width           =   10305
   End
End
Attribute VB_Name = "frmEquipos_Detalle_Verificacion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public PK As Long

Private Sub cmdEtiqueta_Click()
    If Not (lstLista.SelectedItem Is Nothing) Then
        If lstLista.SelectedItem.Index = 1 Then ' sólo si está seleccionada la verificación más actual
            Call imprimir_etiqueta(lstLista.SelectedItem, lstLista.SelectedItem.SubItems(1))
        Else
            MsgBox "Debe estar seleccionada la verificación más actual del equipo" & vbCrLf & _
                   "para generar su etiqueta.", vbOKOnly + vbInformation, App.Title
        End If
    End If
End Sub

Private Sub cmdNuevaVerificacion_Click()
    Call limpiar_datos_verificacion
End Sub

Private Sub Form_Load()
    Dim Titulo As String
    Dim oEQV As New clsEquipos_verificacion
    
    log (Me.Name)
    cargar_botones Me
    Call cargar_combos
    Call cabecera
    
    If oEQV.total_verificaciones(PK) > 0 Then ' si tiene verificaciones
        Call cargar_lista_verificaciones_asignadas
        Call CARGAR_VERIFICACION(listaVerificacionesAsignadas.SelectedItem)   ' se carga
        Call cargar_lista_verificaciones_hco(listaVerificacionesAsignadas.SelectedItem)
    Else
        ' se abre vacío
        Call marcoDatos_Verificacion_activo(False)
    End If
    Set oEQV = Nothing
End Sub

Private Sub listaVerificacionesAsignadas_Click()
    'E0505-I
    If listaVerificacionesAsignadas.ListItems.Count <> 0 Then
        txtDatos(0) = listaVerificacionesAsignadas.SelectedItem
        Call CARGAR_VERIFICACION(listaVerificacionesAsignadas.SelectedItem)
        Call cargar_lista_verificaciones_hco(listaVerificacionesAsignadas.SelectedItem)
    End If
    'E0505-F
End Sub



Private Sub txtdatos_GotFocus(Index As Integer)
    txtDatos(Index).BackColor = &H80C0FF
End Sub

Private Sub txtdatos_LostFocus(Index As Integer)
    txtDatos(Index).BackColor = vbWhite
End Sub

Private Sub cmbPeriVerificacion_Change()
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
        cmbVerificador_interno.Limpiar
        cmbVerificador_interno.Visible = False
        cmbOperador_Interno_Real.Limpiar
        cmbOperador_Interno_Real.Visible = False
        cmbVerificador.Limpiar
        cmbVerificador.cargar_datos
        cmbVerificador.Visible = True
        cmbVerificador.activar
        cmbOperador_Externo_Real.Limpiar
        cmbOperador_Externo_Real.cargar_datos
        cmbOperador_Externo_Real.Visible = True
        cmbOperador_Externo_Real.activar
    ElseIf UCase(cmbTipoVerificacion.Text) = "INTERNA" Then
        cmbVerificador.Limpiar
        cmbVerificador.Visible = False
        cmbOperador_Externo_Real.Limpiar
        cmbOperador_Externo_Real.Visible = False
        cmbVerificador_interno.Limpiar
        cmbVerificador_interno.cargar_datos
        cmbVerificador_interno.Visible = True
        cmbVerificador_interno.activar
        cmbOperador_Interno_Real.Limpiar
        cmbOperador_Interno_Real.cargar_datos
        cmbOperador_Interno_Real.Visible = True
        cmbOperador_Interno_Real.activar
    End If
End Sub

Private Sub cmdModificarVerificacion_Click()
    
    If datos_verificacion_correctos() Then
        Dim oEV As New clsEquipos_verificacion
        Dim lngID_Verificacion As Long
        
        With oEV
            .setEQUIPO_ID = PK
            .setMODALIDAD_ID = cmbTipoVerificacion.BoundText
            .setPERIODICIDAD_ID = cmbPeriVerificacion.BoundText
            .setPROCEDIMIENTO_ID = IIf(cmbProcedimientoVerificacion.BoundText = "", 0, cmbProcedimientoVerificacion.BoundText)
            If UCase(cmbTipoVerificacion.Text) = "INTERNA" Then
                .setVERIFICADOR_INTERNO_ID = cmbVerificador_interno.getPK_SALIDA
                .setVERIFICADOR_EXTERNO_ID = 0
            ElseIf UCase(cmbTipoVerificacion.Text) = "EXTERNA" Then
                .setVERIFICADOR_EXTERNO_ID = cmbVerificador.getPK_SALIDA
                .setVERIFICADOR_INTERNO_ID = 0
            End If
            .setFECHA_ACTUAL = Format(fecha_actual, "yyyy-mm-dd")
            .setFECHA_PROXIMA = Format(fecha_proxima, "yyyy-mm-dd")
            .setRANGO_MIN = txtDatos(1)
            .setRANGO_MAX = txtDatos(2)
            .setUNIDADES_ID = cmbUnidad.getPK_SALIDA
            .setACTIVA = 1
            .setEFECTIVA = 1
        End With
        
        lngID_Verificacion = listaVerificacionesAsignadas.SelectedItem
        If MsgBox("Va a modificar los datos de la verificación. ¿Está seguro?", vbQuestion + vbYesNo, App.Title) = vbYes Then
            oEV.Modificar (lngID_Verificacion)
        Else
            Exit Sub
        End If
        
        'frmEquipos_Detalle.datos_verificacion (PK) ' para actualizar la ventana frmEquipos_Detalle
        Call cargar_lista_verificaciones_asignadas
        MsgBox "La verificación del equipo se ha actualizado correctamente.", vbOKOnly + vbInformation, App.Title
    End If
End Sub

Private Sub cmdGuardar_Verificacion_hco_Click()
    If datos_historico_correctos() Then
        Dim lngEQV_HCO As Long
        Dim oEQV_HCO As New clsEquipos_verificacion_hco
        
        With oEQV_HCO
            .setVERIFICACION_ID = listaVerificacionesAsignadas.SelectedItem
            .setFECHA = Format(Now, "yyyy-mm-dd hh:nn:ss")
            .setMODALIDAD = cmbTipoVerificacion.Text
            .setPROCEDIMIENTO = cmbProcedimientoVerificacion.Text
            If UCase(cmbTipoVerificacion.Text) = "INTERNA" Then
                .setOPERADOR = cmbOperador_Interno_Real.getTEXTO
                .setOPERADOR_ID = cmbOperador_Interno_Real.getPK_SALIDA
            ElseIf UCase(cmbTipoVerificacion.Text) = "EXTERNA" Then
                .setOPERADOR = cmbOperador_Externo_Real.getTEXTO
            End If
            .setLIMITACIONES_USO = txtDatos(4)
            .setRANGO_MIN = txtDatos(1)
            .setRANGO_MAX = txtDatos(2)
            .setUNIDADES = cmbUnidad.getTEXTO
            .setCONFORME = chkConforme.value
        End With
        If MsgBox("Va a introducir una verificación en el histórico. ¿Está seguro?", vbQuestion + vbYesNo, App.Title) = vbYes Then
            lngEQV_HCO = oEQV_HCO.Insertar
            
            Call cargar_lista_verificaciones_hco(listaVerificacionesAsignadas.SelectedItem)
            Call borrar_datos_verificacion_para_hco
            
            fecha_actual = oEQV_HCO.getFECHA
            fecha_proxima = Equipos.calcular_fecha_proxima(fecha_actual.value, cmbPeriVerificacion.Text)
            cmdModificarVerificacion_Click
            
            If MsgBox("La verificación del equipo se insertó correctamente." & vbCrLf & _
                      "¿Desea imprimir la etiqueta de verificación?", vbYesNo + vbInformation, App.Title) = vbYes Then
                Call imprimir_etiqueta(oEQV_HCO.getVERIFICACION_ID, oEQV_HCO.getFECHA)
            End If
        Else
            Exit Sub
        End If
        
        Set oEQV_HCO = Nothing
    End If
End Sub

Private Sub cmdEliminarVerificacion_Click()
    If Not (listaVerificacionesAsignadas.SelectedItem Is Nothing) Then
        If MsgBox("Se va a eliminar la verificación '" & listaVerificacionesAsignadas.SelectedItem.SubItems(1) & "'" & vbCrLf & _
                  "¿Está seguro?", vbYesNo + vbInformation, App.Title) = vbYes Then
                  
            Dim oEV As New clsEquipos_verificacion
            If oEV.Eliminar(listaVerificacionesAsignadas.SelectedItem) Then
                Call cargar_lista_verificaciones_asignadas
                If Not listaVerificacionesAsignadas.SelectedItem Is Nothing Then ' Si el equipo tiene algún plan de verificación
                    Call listaVerificacionesAsignadas_Click
                    Call cargar_lista_verificaciones_hco(listaVerificacionesAsignadas.SelectedItem)
                Else
                    Call limpiar_datos_verificacion
                End If
                MsgBox "La verificación se ha eliminado correctamente.", vbOKOnly + vbInformation, App.Title
            End If
            Set oEV = Nothing
            
            'frmEquipos_Detalle.datos_verificacion (PK) ' para actualizar la ventana frmEquipos_Detalle
            
        End If
    Else
        MsgBox "Debe seleccionar la verificación que desea eliminar.", vbOKOnly + vbInformation, App.Title
    End If
End Sub

Private Sub cmdEliminarVerificacion_Historico_Click()
    If Not (lstLista.SelectedItem Is Nothing) Then
        If MsgBox("Se va a eliminar la verificación realizada el " & lstLista.SelectedItem.SubItems(2) & vbCrLf & _
                  "efectuada por " & lstLista.SelectedItem.SubItems(4) & vbCrLf & _
                  "¿Está seguro?", vbYesNo + vbInformation, App.Title) = vbYes Then
            Dim oEVH As New clsEquipos_verificacion_hco
            If oEVH.Eliminar(lstLista.SelectedItem, lstLista.SelectedItem.SubItems(1)) Then
                Call cargar_lista_verificaciones_hco(listaVerificacionesAsignadas.SelectedItem)
                MsgBox "La verificación se ha eliminado del histórico correctamente.", vbOKOnly + vbInformation, App.Title
            End If
            Set oEVH = Nothing
        End If
    Else
        MsgBox "Debe seleccionar del histórico la verificación que desea eliminar.", vbOKOnly + vbInformation, App.Title
    End If
End Sub

' botón que da de alta una verificación para el equipo
Private Sub cmdok_Click()

    If datos_verificacion_correctos() Then
        Dim oEV As New clsEquipos_verificacion
        Dim lngID_Verificacion As Long
        
        With oEV
            .setEQUIPO_ID = PK
            .setMODALIDAD_ID = cmbTipoVerificacion.BoundText
            .setPERIODICIDAD_ID = cmbPeriVerificacion.BoundText
            .setPROCEDIMIENTO_ID = IIf(cmbProcedimientoVerificacion.BoundText = "", 0, cmbProcedimientoVerificacion.BoundText)
            If UCase(cmbTipoVerificacion.Text) = "INTERNA" Then
                .setVERIFICADOR_INTERNO_ID = cmbVerificador_interno.getPK_SALIDA
                .setVERIFICADOR_EXTERNO_ID = 0
            ElseIf UCase(cmbTipoVerificacion.Text) = "EXTERNA" Then
                .setVERIFICADOR_EXTERNO_ID = cmbVerificador.getPK_SALIDA
                .setVERIFICADOR_INTERNO_ID = 0
            Else
                .setVERIFICADOR_EXTERNO_ID = 0
                .setVERIFICADOR_INTERNO_ID = 0
            End If
            .setFECHA_ACTUAL = Format(fecha_actual, "yyyy-mm-dd")
            .setFECHA_PROXIMA = Format(fecha_proxima, "yyyy-mm-dd")
            .setRANGO_MIN = txtDatos(1)
            .setRANGO_MAX = txtDatos(2)
            .setUNIDADES_ID = cmbUnidad.getPK_SALIDA
            .setACTIVA = 1
            .setEFECTIVA = 0
        End With
        
        If MsgBox("Va a introducir una nueva verificación. ¿Está seguro?", vbQuestion + vbYesNo, App.Title) = vbYes Then
            lngID_Verificacion = oEV.Insertar
        Else
            Exit Sub
        End If
        
        Call cargar_lista_verificaciones_asignadas
        Call seleccionar_verificacion_de_asignadas(Trim(txtDatos(0)))
        
        'frmEquipos_Detalle.datos_verificacion (PK) ' para actualizar la ventana frmEquipos_Detalle
        estado_botones (True)
        MsgBox "La verificación del equipo se ha creado correctamente.", vbOKOnly + vbInformation, App.Title
    End If
End Sub

Private Sub cmdcancel_Click()
    Unload Me
End Sub

' ----------------- Funciones auxiliares del formulario ----------------
Public Sub cabecera()
    ' lista verificaciones asignadas
    With listaVerificacionesAsignadas.ColumnHeaders
        .Add , , "ID", 500, lvwColumnLeft
        .Add , , "Modalidad", 2000, lvwColumnLeft
        .Add , , "Periodicidad", 2000, lvwColumnLeft
        .Add , , "Responsable", 2500, lvwColumnLeft
    End With
    
    ' lista verificaciones historico
    With lstLista.ColumnHeaders
        .Add , , "ID", 0, lvwColumnLeft
        .Add , , "Fecha hora", 0, lvwColumnLeft
        .Add , , "Fecha", 1200, lvwColumnLeft
        .Add , , "Modalidad", 1000, lvwColumnLeft
        .Add , , "Realizado por", 2000, lvwColumnLeft
        .Add , , "Limitaciones uso", 1500, lvwColumnLeft
        .Add , , "Rango Min", 1000, lvwColumnLeft
        .Add , , "Rango Máx", 1000, lvwColumnLeft
        .Add , , "Unidades", 1000, lvwColumnLeft
        .Add , , "Conforme", 1000, lvwColumnLeft
        .Add , , "OPERADOR_ID", 0, lvwColumnLeft
    End With
End Sub

Private Sub cargar_combos()
    Dim oDECO As New clsDecodificadora
    
    oDECO.Cargar_Combo cmbTipoVerificacion, decodificadora.EQ_TIPO_CALIBRACION
    oDECO.Cargar_Combo cmbPeriVerificacion, decodificadora.EQ_periodicidad
    llenar_combo cmbVerificador, New clsProveedor, 0, frmProveedores, ""
    llenar_combo cmbVerificador_interno, New clsUsuarios, 0, Me, ""
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

' función que carga el formulario con los datos de la verificación pasada por parámetro
Public Sub CARGAR_VERIFICACION(lngVerificacion_id As Long)
    Dim oEquipo As New clsEquipos
    
    If oEquipo.Carga(PK) = True Then
        lbltitulo = "Verificación del Equipo : " & oEquipo.getNOMBRE
        Me.Caption = lbltitulo
        Dim oEV As New clsEquipos_verificacion
        If oEV.Carga(lngVerificacion_id) Then
            With oEV
                cmbTipoVerificacion.BoundText = .getMODALIDAD_ID
                cmbVerificador.MostrarElemento .getVERIFICADOR_EXTERNO_ID
                cmbVerificador_interno.MostrarElemento .getVERIFICADOR_INTERNO_ID
                
                cmbOperador_Externo_Real.MostrarElemento .getVERIFICADOR_EXTERNO_ID
                cmbOperador_Interno_Real.MostrarElemento .getVERIFICADOR_INTERNO_ID
                
                cmbPeriVerificacion.BoundText = .getPERIODICIDAD_ID
                cmbProcedimientoVerificacion.BoundText = .getPROCEDIMIENTO_ID
                fecha_actual = .getFECHA_ACTUAL
                fecha_proxima = .getFECHA_PROXIMA
                txtDatos(1) = .getRANGO_MIN
                txtDatos(2) = .getRANGO_MAX
                cmbUnidad.MostrarElemento .getUNIDADES_ID
                
                Set oEV = Nothing
            End With
        End If
    End If
    Set oEquipo = Nothing
End Sub

' función que carga la lista de verificaciones asignadas al equipo
Private Sub cargar_lista_verificaciones_asignadas()
    Dim rs As ADODB.RecordSet
    Dim oEV As New clsEquipos_verificacion
    
    listaVerificacionesAsignadas.ListItems.Clear
    Set rs = oEV.Listado_verificaciones_asignadas(PK)
    If rs.RecordCount <> 0 Then
        Do
            With listaVerificacionesAsignadas.ListItems.Add(, , Format(rs(0), "0000"))
                .SubItems(1) = rs(1)
                .SubItems(2) = rs(2)
                .SubItems(3) = rs(3)
            End With
            rs.MoveNext
        Loop Until rs.EOF
        cmdModificarVerificacion.Enabled = True
    Else ' si el equipo no tiene ninguna verificación asignada se deshabilita el botón 'modificar'
        cmdModificarVerificacion.Enabled = False
    End If
    
    Set oEV = Nothing
End Sub

' procedimiento que selecciona la verificación con el nombre pasado
' por parámetro de la lista de verificaciones asignadas
Private Sub seleccionar_verificacion_de_asignadas(strId As String)
    Dim i As Long
    
    For i = 1 To listaVerificacionesAsignadas.ListItems.Count
        If listaVerificacionesAsignadas.ListItems(i) = strId Then
            listaVerificacionesAsignadas.ListItems(i).Selected = True
        End If
    Next i
End Sub

' Procedimiento que carga el histórico de verificaciones
' de la verificación pasada por parámetro
Private Sub cargar_lista_verificaciones_hco(lngVerificacion As Long)
    Dim rs As ADODB.RecordSet
    Dim oEVH As New clsEquipos_verificacion_hco
    
    lstLista.ListItems.Clear
    Set rs = oEVH.Listado(lngVerificacion)
    If rs.RecordCount <> 0 Then
        Do
            With lstLista.ListItems.Add(, , Format(rs(0), "0000"))
                .SubItems(1) = Format(rs(1), "yyyy-mm-dd hh:nn:ss")
                .SubItems(2) = Format(rs(2), "dd/mm/yyyy")
                .SubItems(3) = rs(3)
                .SubItems(4) = rs(4)
                .SubItems(5) = rs(5)
                .SubItems(6) = rs(6)
                .SubItems(7) = rs(7)
                .SubItems(8) = rs(8)
                .SubItems(9) = rs(9)
                .SubItems(10) = rs(10)
            End With
            rs.MoveNext
        Loop Until rs.EOF
    End If
    Set oEVH = Nothing
End Sub

Private Sub marcoDatos_Verificacion_activo(booEstado As Boolean)
    frmDatosVerificacion.Enabled = booEstado
    cmbOperador_Externo_Real.desactivar
    cmbOperador_Interno_Real.desactivar
    chkConforme.Enabled = False
    cmdGuardar_Verificacion_hco.Enabled = booEstado
    cmdModificarVerificacion.Enabled = booEstado
End Sub

' función que comprueba que los datos que se van a introducir
' para el histórico son correctos.
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
    
    If Trim(txtDatos(4)) = "" Then ' limitaciones de uso
        MsgBox "Debe indicar limitaciones de uso de la verificación.", vbInformation, App.Title
        txtDatos(4).SetFocus
        datos_historico_correctos = False
        Exit Function
    End If
    
    If Trim(cmbPeriVerificacion) = "" Then
        MsgBox "Debe indicar la periodicidad de la verificación.", vbInformation, App.Title
        cmbPeriVerificacion.SetFocus
        datos_historico_correctos = False
        Exit Function
    End If

    datos_historico_correctos = True
End Function

Private Sub borrar_datos_verificacion_para_hco()
    cmbOperador_Interno_Real.Limpiar
    cmbOperador_Externo_Real.Limpiar
    txtDatos(4).Text = ""
    chkConforme.value = 0
End Sub

Private Sub limpiar_datos_verificacion()
    Call borrar_datos_verificacion_para_hco
    
    txtDatos(1).Text = ""
    txtDatos(2).Text = ""
    cmbTipoVerificacion.Text = ""
    cmbPeriVerificacion.Text = ""
    cmbProcedimientoVerificacion.Text = ""
    cmbVerificador.Limpiar
    cmbVerificador_interno.Limpiar
    fecha_actual = "1900-01-01"
    fecha_proxima = "1900-01-01"
    cmbUnidad.Limpiar
    estado_botones (False)
    lstLista.ListItems.Clear ' se borra la lista de histórico
    
End Sub

Private Sub estado_botones(booEstado As Boolean)
    listaVerificacionesAsignadas.Enabled = booEstado
    cmdModificarVerificacion.Enabled = booEstado
    cmdGuardar_Verificacion_hco.Enabled = booEstado
End Sub

' función que comprueba que los datos de la verificación
' para su introducción son correctos.
Private Function datos_verificacion_correctos() As Boolean
    
    datos_verificacion_correctos = True
    
    If Trim(cmbTipoVerificacion) = "" Then ' modalidad
        MsgBox "Debe seleccionar una modalidad.", vbInformation, App.Title
        cmbTipoVerificacion.SetFocus
        datos_verificacion_correctos = False
        Exit Function
    End If
    
    If UCase(cmbTipoVerificacion) = "INTERNA" Then ' responsable
        If Trim(cmbVerificador_interno.getTEXTO) = "" Then
            MsgBox "Debe seleccionar un responsable.", vbInformation, App.Title
            cmbVerificador_interno.SetFocus
            datos_verificacion_correctos = False
            Exit Function
        End If
    ElseIf UCase(cmbTipoVerificacion) = "EXTERNA" Then
        If Trim(cmbVerificador.getTEXTO) = "" Then
            MsgBox "Debe seleccionar un responsable.", vbInformation, App.Title
            cmbVerificador.SetFocus
            datos_verificacion_correctos = False
            Exit Function
        End If
    End If
    
    If Trim(cmbPeriVerificacion) = "" Then ' periodicidad
        MsgBox "Debe seleccionar una periodicidad.", vbInformation, App.Title
        cmbPeriVerificacion.SetFocus
        datos_verificacion_correctos = False
        Exit Function
    End If
    
End Function

Private Sub imprimir_etiqueta(strID_Verificacion As Long, strFecha_Verificacion As String)
    Dim prnPrinter As Printer
    
On Error GoTo trataError

    ' se mira si el equipo tiene impresora de etiquetas
    Dim oParametro As New clsParametros
    If Not oParametro.Carga(parametros.IMPRESORA_ETIQUETAS_GRANDE, USUARIO.getUSO) Then
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
    
    With frmReport
        Firmas.copiar_firma_responsable_tecnico
        .iniciar
        .informe = "rptEquipos_ETIQUETA_Verificacion"
        .criterio = "{equipos.ID_EQUIPO} = " & CLng(PK) & _
                    " and {equipos_verificacion_historico.VERIFICACION_ID} = " & strID_Verificacion & _
                    " and {equipos_verificacion_historico.FECHA} = datetime('" & strFecha_Verificacion & "')"
        .imprimir = True
        .generar
        .Visible = False
    End With
    log ("Final impresion de etiqueta de verificación de equipo")
    
    Exit Sub
    
trataError:
    MsgBox "Error al imprimir la etiqueta de verificación.", vbCritical, Err.Description
End Sub
