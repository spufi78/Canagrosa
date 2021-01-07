VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{CEA3B7F6-2847-4E5E-A551-DB7A62489D44}#46.0#0"; "miCombo.ocx"
Begin VB.Form frmEquipoEdicionMtoFechas 
   Appearance      =   0  'Flat
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Generar Fechas Mantenimiento Equipo"
   ClientHeight    =   8445
   ClientLeft      =   2445
   ClientTop       =   1680
   ClientWidth     =   12555
   Icon            =   "frmEquipoEdicionMtoFechas.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8445
   ScaleWidth      =   12555
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdok 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Aceptar"
      Height          =   870
      Left            =   10380
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   7530
      Width           =   1050
   End
   Begin VB.CommandButton cmdcancel 
      BackColor       =   &H00E0E0E0&
      Caption         =   "ESC-Salir"
      Height          =   870
      Left            =   11460
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   7530
      Width           =   1050
   End
   Begin VB.Frame frmDatosEquipo 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Datos del Equipo"
      Height          =   3405
      Left            =   0
      TabIndex        =   2
      Top             =   600
      Width           =   12525
      Begin MSComctlLib.ListView lstPlanes 
         Height          =   2175
         Left            =   60
         TabIndex        =   25
         Top             =   1200
         Width           =   7365
         _ExtentX        =   12991
         _ExtentY        =   3836
         View            =   3
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         Checkboxes      =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   12640511
         BorderStyle     =   1
         Appearance      =   0
         NumItems        =   0
      End
      Begin VB.CheckBox chkAPartirDeFecha 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Generar a partir de un día concreto (inclusive)"
         Height          =   255
         Left            =   7500
         TabIndex        =   24
         Top             =   1890
         Width           =   4125
      End
      Begin VB.CommandButton cmdGenerarPlan 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Generar Plan"
         Height          =   870
         Left            =   11280
         Style           =   1  'Graphical
         TabIndex        =   19
         Top             =   2490
         Width           =   1170
      End
      Begin VB.ComboBox cmbDia 
         Height          =   315
         ItemData        =   "frmEquipoEdicionMtoFechas.frx":1272
         Left            =   7500
         List            =   "frmEquipoEdicionMtoFechas.frx":12E9
         Style           =   2  'Dropdown List
         TabIndex        =   18
         Top             =   2190
         Visible         =   0   'False
         Width           =   1035
      End
      Begin VB.ComboBox cmbMes 
         Height          =   315
         ItemData        =   "frmEquipoEdicionMtoFechas.frx":1360
         Left            =   8580
         List            =   "frmEquipoEdicionMtoFechas.frx":138B
         Style           =   2  'Dropdown List
         TabIndex        =   17
         Top             =   2190
         Visible         =   0   'False
         Width           =   1725
      End
      Begin VB.ComboBox cmdAnoReferencia 
         Height          =   315
         ItemData        =   "frmEquipoEdicionMtoFechas.frx":13F4
         Left            =   8730
         List            =   "frmEquipoEdicionMtoFechas.frx":1473
         Style           =   2  'Dropdown List
         TabIndex        =   16
         Top             =   1560
         Width           =   1035
      End
      Begin VB.TextBox txtFamilia 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         Height          =   315
         Left            =   8160
         Locked          =   -1  'True
         TabIndex        =   11
         Top             =   570
         Width           =   4305
      End
      Begin VB.TextBox txtNSerie 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         Height          =   315
         Left            =   8160
         Locked          =   -1  'True
         TabIndex        =   9
         Top             =   210
         Width           =   4305
      End
      Begin VB.TextBox txtModelo 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         Height          =   315
         Left            =   1200
         Locked          =   -1  'True
         TabIndex        =   7
         Top             =   570
         Width           =   4995
      End
      Begin VB.TextBox txtNombre 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         Height          =   315
         Left            =   1200
         Locked          =   -1  'True
         TabIndex        =   5
         Top             =   210
         Width           =   4995
      End
      Begin pryCombo.miCombo cmbResponsable 
         Height          =   315
         Left            =   7500
         TabIndex        =   22
         Top             =   1200
         Width           =   4965
         _ExtentX        =   8758
         _ExtentY        =   556
      End
      Begin MSDataListLib.DataCombo cmbProcedimiento 
         Height          =   315
         Left            =   11340
         TabIndex        =   23
         Top             =   1590
         Visible         =   0   'False
         Width           =   1035
         _ExtentX        =   1826
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         Style           =   2
         Text            =   ""
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "(Se ha de señalar al menos un para generar las fechas)"
         ForeColor       =   &H00FF0000&
         Height          =   195
         Index           =   4
         Left            =   3480
         TabIndex        =   26
         Top             =   990
         Width           =   3900
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Responsable"
         Height          =   195
         Index           =   2
         Left            =   7500
         TabIndex        =   21
         Top             =   960
         Width           =   930
      End
      Begin VB.Label lblAviso 
         BackColor       =   &H00C0C0C0&
         Caption         =   $"frmEquipoEdicionMtoFechas.frx":14F5
         ForeColor       =   &H000000C0&
         Height          =   795
         Left            =   7500
         TabIndex        =   20
         Top             =   2580
         Visible         =   0   'False
         Width           =   3585
      End
      Begin VB.Label Label1 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Año Referencia"
         Height          =   195
         Left            =   7500
         TabIndex        =   15
         Top             =   1650
         Width           =   1155
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Familia"
         Height          =   195
         Index           =   5
         Left            =   7500
         TabIndex        =   10
         Top             =   630
         Width           =   480
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Nº Serie"
         Height          =   195
         Index           =   3
         Left            =   7500
         TabIndex        =   8
         Top             =   300
         Width           =   585
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Modelo"
         Height          =   195
         Index           =   1
         Left            =   90
         TabIndex        =   6
         Top             =   630
         Width           =   525
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Equipo"
         Height          =   195
         Index           =   0
         Left            =   120
         TabIndex        =   4
         Top             =   330
         Width           =   495
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Planes de Mantenimiento"
         Height          =   195
         Index           =   10
         Left            =   120
         TabIndex        =   3
         Top             =   990
         Width           =   1785
      End
   End
   Begin MSFlexGridLib.MSFlexGrid grdFechas 
      Height          =   3435
      Left            =   30
      TabIndex        =   12
      Top             =   4050
      Width           =   12495
      _ExtentX        =   22040
      _ExtentY        =   6059
      _Version        =   393216
      Rows            =   1
      Cols            =   5
      FixedCols       =   0
      BackColor       =   12640511
      BackColorSel    =   16576
      FocusRect       =   2
      HighLight       =   2
      FillStyle       =   1
      GridLinesFixed  =   1
      SelectionMode   =   1
      AllowUserResizing=   1
      Appearance      =   0
   End
   Begin VB.Label lblsubtitulo 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "Generar Fechas para el Mantenimiento de Equipos"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   90
      TabIndex        =   1
      Top             =   315
      Width           =   3585
   End
   Begin VB.Image imagen 
      Height          =   480
      Left            =   11970
      Picture         =   "frmEquipoEdicionMtoFechas.frx":1595
      Top             =   45
      Width           =   480
   End
   Begin VB.Label lbltitulo 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "Mantenimiento de Equipo"
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
      TabIndex        =   0
      Top             =   45
      Width           =   2640
   End
   Begin VB.Shape fondo 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   600
      Left            =   0
      Top             =   0
      Width           =   12555
   End
End
Attribute VB_Name = "frmEquipoEdicionMtoFechas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private mvarblnResultado As Boolean
Private mvarobjEquipo As clsEquipos
Private mvarobjPlanMantenimiento As clsPlanMantenimiento
Private mvarobjFechasMtoEquipo As New clsGenericCollection
Private mvarlngidResponsable As Long
Private mvarblnVieneDeCuaderno As Boolean
Private mvarlngPK As Long
Private mvarlngidEquipo As Long
Private mvardtmFechaPrevista As Date
Private mvarlngIdEvento As Long

Private oFechasMto As New clsEquipoMantenimiento

Private mvar_lista_planes As String
Private Sub cabecera()

        With lstPlanes.ColumnHeaders
            .Add , , "Plan Mto", 3680, lvwColumnLeft
            .Add , , "Protocolo", 3680, lvwColumnLeft
            .Add , , "id_plan", 0, lvwColumnLeft
            .Add , , "id_protocolo", 0, lvwColumnLeft
    End With
    
End Sub

Public Property Get idEquipo() As Long

    idEquipo = mvarlngidEquipo

End Property

Public Property Let idEquipo(ByVal lngidEquipo As Long)

    mvarlngidEquipo = lngidEquipo

End Property


Public Property Get VieneDeCuaderno() As Boolean

    VieneDeCuaderno = mvarblnVieneDeCuaderno

End Property

Public Property Let VieneDeCuaderno(ByVal blnVieneDeCuaderno As Boolean)

    mvarblnVieneDeCuaderno = blnVieneDeCuaderno

End Property

Public Property Get idResponsable() As Long

    idResponsable = mvarlngidResponsable

End Property

Public Property Let idResponsable(ByVal lngidResponsable As Long)

    mvarlngidResponsable = lngidResponsable

End Property

Private Sub cmdcancel_Click()
    mvarblnResultado = False
    Me.Hide
End Sub

Private Sub cmdGenerarPlan_Click()
Dim intAnno As Integer, lngidResponsable As Long, lngidProcedimiento As Long
Dim strProcedimiento As String, strRESPONSABLE As String
Dim dtmFechaAPartir As Date
Dim objItem As clsEquipoMantenimiento
Dim objCol As New clsGenericCollection
objCol.KeyName = "setID_MANTENIMIENTO"
Dim sql As String
Dim oMantenimiento As New clsEquipoMantenimiento
Dim oPM As clsPlanMantenimiento
Dim lngIdAcumulado As Long

    lngIdAcumulado = 0
    
    If Not comprobar_planes_seleccionados Then
        Set oMantenimiento = Nothing
        Set objItem = Nothing
        Set objCol = Nothing
        Set oPM = Nothing
        Exit Sub
    End If

    If mvarobjFechasMtoEquipo.Count > 0 Then
        If MsgBox("ATENCIÓN: Ha generado fechas de Mantenimiento. Si vuelve a generar, eliminará las existentes. ¿Desea Continuar?", vbInformation + vbYesNo, "Generar Fechas de Mantenimiento") = vbNo Then
            Exit Sub
        End If
    End If

    If cmdAnoReferencia.ListIndex < 0 Then
        MsgBox "Debe Señalar un Año para generar Fechas", vbInformation, "Generar Fechas de Mantenimiento"
        Exit Sub
    End If
    
    intAnno = CInt(cmdAnoReferencia.List(cmdAnoReferencia.ListIndex))
    
    If chkAPartirDeFecha.Value = vbChecked Then
        If cmbDia.ListIndex < 0 Then
            MsgBox "Si desea generar las fechas a partir de un día en concreto, debe señalar el Día y Mes a partir del cual generar dichas fechas", vbInformation, "Generar Fechas de Mantenimiento"
            Exit Sub
        ElseIf cmbMes.ListIndex < 0 Then
            MsgBox "Si desea generar las fechas a partir de un día en concreto, debe señalar el Día y Mes a partir del cual generar dichas fechas", vbInformation, "Generar Fechas de Mantenimiento"
            Exit Sub
        Else
            dtmFechaAPartir = DateSerial(intAnno, CInt(cmbMes.ItemData(cmbMes.ListIndex)), CInt(cmbDia.List(cmbDia.ListIndex)))
        End If
    End If
    
'    If getDataComboSel(cmbProcedimiento) <= 0 Then
'        MsgBox "Debe Señalar un Procedimiento de Mantenimiento para generar Fechas", vbInformation, "Generar Fechas de Mantenimiento"
'        Exit Sub
'    End If
'    lngidProcedimiento = getDataComboSel(cmbProcedimiento)
'    strProcedimiento = cmbProcedimiento.Text
'
    
'    If cmbProtocolo.getPK_SALIDA <= 0 Then
'        MsgBox "Debe Señalar un Protocolo de Mantenimiento para generar Fechas", vbInformation, "Generar Fechas de Mantenimiento"
'        Exit Sub
'    End If
'    lngidProcedimiento = cmbProtocolo.getPK_SALIDA
'    strProcedimiento = cmbProtocolo.getTEXTO
    
    
    
    If cmbResponsable.getPK_SALIDA <= 0 Then
        MsgBox "Debe Señalar un Responsable de Mantenimiento para generar Fechas", vbInformation, "Generar Fechas de Mantenimiento"
        Exit Sub
    End If
    lngidResponsable = cmbResponsable.getPK_SALIDA
    strRESPONSABLE = cmbResponsable.getTEXTO
    
    
    ' a partir de aqui, el bucle para cada mantenimiento
    'NOTA: POR CADA ITERACION, CAMBIA EL PLAN.
    Dim x As Integer
    Set mvarobjFechasMtoEquipo = New clsGenericCollection
    mvarobjFechasMtoEquipo.KeyName = "setID_MANTENIMIENTO"
    
    lngIdAcumulado = -1
    
    For x = 1 To lstPlanes.ListItems.Count
        If lstPlanes.ListItems(x).Checked Then
            mvar_lista_planes = mvar_lista_planes & ":" & lstPlanes.ListItems(x).SubItems(2) & ":"
            Set oPM = New clsPlanMantenimiento
            oPM.Carga CLng(lstPlanes.ListItems(x).SubItems(2)) ' carga el Plan de mantenimiento
            
            Set objCol = oPM.generarFechasPlanMto(intAnno)
            
            For Each objItem In objCol.Iterator
                ' Le añade los datos que le falten
                objItem.setEQUIPO_ID = mvarobjEquipo.getID_EQUIPO
                objItem.setMANTENEDOR_ID = lngidResponsable
                objItem.setPLANMTO_ID = oPM.getID_PLAN_MTO
                objItem.setPLAN_MANTENIMIENTO = oPM.getNOMBRE ' & ": " & oPM.getDESCRIPCION
                objItem.setRESPONSABLE = strRESPONSABLE
                objItem.setPROCEDIMIENTO_ID = oPM.getPROTOCOLO_ID ' lngidProcedimiento
                objItem.setPROCEDIMIENTO = oPM.getPROTOCOLO ' strProcedimiento
                
                sql = "INSERT INTO eq_mantenimiento_equipos (ID_MANTENIMIENTO, EQUIPO_ID, PLANMTO_ID, PROCEDIMIENTO_ID, MANTENEDOR_ID, FECHA_ACTUAL, OBSERVACIONES, ESTADO, CUSERID, MUSERID, TS, RUTA_CERTIFICADO)"
                sql = sql & "SELECT COALESCE(MAX(ID_MANTENIMIENTO), 0) + 1 AS ID_MANTENIMIENTO"
                sql = sql & ", " & objItem.getEQUIPO_ID
                sql = sql & ", " & objItem.getPLANMTO_ID
                sql = sql & ", " & objItem.getPROCEDIMIENTO_ID
                sql = sql & ", " & objItem.getMANTENEDOR_ID
                sql = sql & ", '" & Format(CDate(objItem.getFECHA_ACTUAL), "yyyy/mm/dd") & "'"
                sql = sql & ", '', 0, " & USUARIO.getID_EMPLEADO
                sql = sql & ", " & USUARIO.getID_EMPLEADO
                sql = sql & ", LOCALTIMESTAMP, ''"
                sql = sql & " FROM eq_mantenimiento_equipos"
                
                
                
                If chkAPartirDeFecha.Value = vbUnchecked Then
                    Call mvarobjFechasMtoEquipo.Add(objItem, CStr(lngIdAcumulado))
                    oMantenimiento.SQLGenesisFechasMto = sql
                    lngIdAcumulado = lngIdAcumulado - 1
                Else
                    If CDate(objItem.getFECHA_ACTUAL) >= dtmFechaAPartir Then
                        Call mvarobjFechasMtoEquipo.Add(objItem, CStr(lngIdAcumulado))
                        oMantenimiento.SQLGenesisFechasMto = sql
                        lngIdAcumulado = lngIdAcumulado - 1
                    End If
                End If
                        
            Next objItem
        End If
    Next x

    Set oFechasMto = oMantenimiento

    Call PresentarDatos_FechasMantenimiento

End Sub

Private Sub cmdok_Click()

    Dim oOP As New clsEquiposOperacionesPendientes
    
    oFechasMto.InsertarFechasMtoGeneradas mvarobjEquipo.getID_EQUIPO
    
    oOP.crear_mantenimientos_pendiente_para_nuevas_fechas_planes mvarobjEquipo.getID_EQUIPO, mvar_lista_planes
        
    mvarlngidResponsable = cmbResponsable.getPK_SALIDA

    mvarblnResultado = True
    Me.Hide
End Sub

Private Sub chkAPartirDeFecha_Click()

cmbDia.visible = (chkAPartirDeFecha.Value = vbChecked)
cmbMes.visible = (chkAPartirDeFecha.Value = vbChecked)

End Sub

Private Sub Form_Load()
    
    'If mvarblnVieneDeCuaderno Then
    '    Call mvarobjEquipo.Carga(mvarlngidEquipo)
    '    mvarenuTipoEdicion = ALTA
    'End If
    log Me.Name
    cargar_botones Me
    
    cargar_combos
    
    cabecera
    
    Call ConfigurarCombo
    
    mvarobjFechasMtoEquipo.KeyName = "setID_MANTENIMIENTO"
    
    'If mvarobjPlanMantenimiento Is Nothing Then
    '    MsgBox ""
    'End If
    
    Call PresentarDatos
    
End Sub

Private Sub Form_Unload(Cancel As Integer)

    Set mvarobjEquipo = Nothing

    Set mvarobjPlanMantenimiento = Nothing
    Set mvarobjFechasMtoEquipo = Nothing
End Sub


Public Property Get RESULTADO() As Boolean

    RESULTADO = mvarblnResultado

End Property

Public Property Let RESULTADO(ByVal blnResultado As Boolean)

    mvarblnResultado = blnResultado

End Property

Public Property Get EQUIPO() As clsEquipos

    Set EQUIPO = mvarobjEquipo

End Property

Public Property Set EQUIPO(objEquipo As clsEquipos)

    Set mvarobjEquipo = objEquipo

End Property

Public Property Get PlanMantenimiento() As clsPlanMantenimiento

    Set PlanMantenimiento = mvarobjPlanMantenimiento

End Property

Public Property Set PlanMantenimiento(objPlanMantenimiento As clsPlanMantenimiento)

    Set mvarobjPlanMantenimiento = objPlanMantenimiento

End Property

Public Property Get FechasMtoEquipo() As clsGenericCollection

    Set FechasMtoEquipo = mvarobjFechasMtoEquipo

End Property

Public Property Set FechasMtoEquipo(objFechasMtoEquipo As clsGenericCollection)

    Set mvarobjFechasMtoEquipo = objFechasMtoEquipo

End Property

Private Sub ConfigurarCombo()
    With grdFechas
        .ColWidth(0) = 0
        .ColWidth(1) = .Width * 0.1
        .ColWidth(2) = .Width * 0.2
        .ColWidth(3) = .Width * 0.4
        .ColWidth(4) = .Width * 0.3
        
        .TextMatrix(0, 1) = "Fecha Prevista"
        .TextMatrix(0, 2) = "Responsable"
        .TextMatrix(0, 3) = "Plan Mto."
        .TextMatrix(0, 4) = "Protocolo"
        
    End With

   


End Sub
Private Sub cargar_combos()
Dim oca_doc As New clsCa_documentos
    
    
    llenar_combo cmbResponsable, New clsUsuarios, 0, Me, ""
    'llenar_combo cmbProtocolo, New clsCa_documentos, 0, frmCA_Documento, ""
    
    
    'Set cmbProcedimiento.RowSource = oCA_Doc.Listado_Combo_procedimientos_calibracion()
    'cmbProcedimiento.ListField = "nombre" 'campo que veo
    'cmbProcedimiento.DataField = "id" 'campo asociado
    'cmbProcedimiento.BoundColumn = "id_documento" 'lo que realmente envia
    'Set oCA_Doc = Nothing
    

    
    
End Sub

Private Sub PresentarDatos()

    PresentarDatos_PlanesMantenimiento


    ' Datos del Equipo
    txtNombre.Text = mvarobjEquipo.getNOMBRE
    txtModelo.Text = mvarobjEquipo.getMODELO
    txtNSerie.Text = mvarobjEquipo.getSERIE
    txtFamilia.Text = mvarobjEquipo.getFAMILIA.getNOMBRE

    ' Responsable y Procedimiento
    Call cmbResponsable.MostrarElemento(mvarlngidResponsable)


    'Call PresentarDatos_PlanMto
    

End Sub

Private Sub PresentarDatos_PlanMto()
'Dim objItem As clsPlanMantenimientoAcciones
'Dim rs_acciones As New ADODB.RecordSet
    
Exit Sub
    
'    txtPlan.Text = mvarobjPlanMantenimiento.getNOMBRE
'    txtPeriodicidad.Text = mvarobjPlanMantenimiento.getFRECUENCIA
'    cmbProtocolo.MostrarElemento mvarobjPlanMantenimiento.getPROTOCOLO_ID
'    lstAcciones.Clear
'
'    Set rs_acciones = mvarobjPlanMantenimiento.devolver_acciones(mvarobjPlanMantenimiento.getID_PLAN_MTO)
'
'    If rs_acciones.RecordCount <> 0 Then
'        rs_acciones.MoveFirst
'
'        While Not rs_acciones.EOF
'            lstAcciones.AddItem (rs_acciones("Nombre"))
'            rs_acciones.MoveNext
'        Wend
'
'    End If
'
'    Set rs_acciones = Nothing
''    For Each objItem In mvarobjPlanMantenimiento.Acciones
''        lstAcciones.AddItem objItem.getNOMBRE
''        lstAcciones.ItemData(lstAcciones.ListCount - 1) = objItem.getID_ACCION
''    Next objItem
'

End Sub

Private Sub PresentarDatos_FechasMantenimiento()
Dim objItem As clsEquipoMantenimiento

    With grdFechas
        .Rows = 1
        If mvarobjFechasMtoEquipo.Count = 0 Then
            MsgBox "No se ha generado ninguna fecha, según este Plan de Manteniento de Equipo", vbInformation, "Generar Fechas de Mantenimiento"
            Exit Sub
        End If
        
        For Each objItem In mvarobjFechasMtoEquipo.Iterator
            .Rows = .Rows + 1
            .TextMatrix(.Rows - 1, 0) = CStr(objItem.getID_MANTENIMIENTO)
            .TextMatrix(.Rows - 1, 1) = objItem.getFECHA_ACTUAL
            .TextMatrix(.Rows - 1, 2) = objItem.getRESPONSABLE
            .TextMatrix(.Rows - 1, 3) = objItem.getPLAN_MANTENIMIENTO
            .TextMatrix(.Rows - 1, 4) = objItem.getPROCEDIMIENTO
        Next
    End With

End Sub



Public Property Get PK() As Long

    PK = mvarlngPK

End Property

Public Property Let PK(ByVal lngPK As Long)

    mvarlngPK = lngPK

End Property



Public Property Get FechaPrevista() As Date

    FechaPrevista = mvardtmFechaPrevista

End Property

Public Property Let FechaPrevista(ByVal dtmFechaPrevista As Date)

    mvardtmFechaPrevista = dtmFechaPrevista

End Property

Public Property Get IdEvento() As Long

    IdEvento = mvarlngIdEvento

End Property

Public Property Let IdEvento(ByVal lngIdEvento As Long)

    mvarlngIdEvento = lngIdEvento

End Property

Private Sub PresentarDatos_PlanesMantenimiento()

    Dim rs As ADODB.Recordset
        
    Set rs = mvarobjEquipo.devolver_lista_planes_mantenimiento
    
    lstPlanes.ListItems.Clear
    
    If rs.RecordCount = 0 Then Exit Sub
    
    rs.MoveFirst
    While Not rs.EOF
        With lstPlanes.ListItems.Add(, , rs("NOMBRE"))
            .SubItems(1) = rs("protocolo")
            .SubItems(2) = rs("plan_mantenimiento_id")
            .SubItems(3) = rs("PROTOCOLO_ID")
        End With
        rs.MoveNext
    Wend
    

End Sub

Private Function comprobar_planes_seleccionados() As Boolean

    Dim intCont As Integer, x As Integer
    
    comprobar_planes_seleccionados = False
    intCont = 0
    
    If lstPlanes.ListItems.Count > 0 Then
        For x = 1 To lstPlanes.ListItems.Count
            If lstPlanes.ListItems(x).Checked Then
                intCont = intCont + 1
            End If
        Next x
    End If
    
    If intCont = 0 Then
        MsgBox "Es necesario señalar al menos un Plan de Mantenimiento para generar las Fechas", vbInformation, "Generar Fechas de Mantenimiento"
        Exit Function
    End If

    comprobar_planes_seleccionados = True
    
End Function
