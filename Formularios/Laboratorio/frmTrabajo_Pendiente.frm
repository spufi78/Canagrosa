VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{CEA3B7F6-2847-4E5E-A551-DB7A62489D44}#46.0#0"; "miCombo.ocx"
Begin VB.Form frmTrabajo_Pendiente 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Listado de Tareas"
   ClientHeight    =   9045
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   13185
   Icon            =   "frmTrabajo_Pendiente.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   9045
   ScaleWidth      =   13185
   Begin VB.CommandButton cmdDeter 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Registro"
      Height          =   870
      Left            =   90
      Picture         =   "frmTrabajo_Pendiente.frx":0442
      Style           =   1  'Graphical
      TabIndex        =   21
      Top             =   8100
      Width           =   1095
   End
   Begin VB.CommandButton cmdEtiqueta 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Etiqueta"
      Height          =   870
      Left            =   1215
      Style           =   1  'Graphical
      TabIndex        =   20
      ToolTipText     =   "Muestra el informe de registrro de la muestra"
      Top             =   8100
      Width           =   1095
   End
   Begin VB.CommandButton cmdListado 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Generar Listado"
      Height          =   870
      Left            =   9945
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   8115
      Width           =   2085
   End
   Begin VB.CommandButton cmdsalir 
      BackColor       =   &H00E0E0E0&
      Caption         =   "ESC-Salir"
      Height          =   870
      Left            =   12060
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   8115
      Width           =   1050
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Datos de Búsqueda"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1425
      Left            =   60
      TabIndex        =   0
      Top             =   810
      Width           =   13095
      Begin VB.CheckBox chkMias 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Mostrar sólo las mías"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   285
         Left            =   225
         TabIndex        =   22
         Top             =   1080
         Width           =   2745
      End
      Begin VB.CheckBox chkcerradas 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Mostrar muestras cerradas"
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
         Left            =   9090
         TabIndex        =   15
         Top             =   -45
         Visible         =   0   'False
         Width           =   2745
      End
      Begin VB.CheckBox chkTodas 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Caption         =   "Todas"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   9660
         TabIndex        =   10
         Top             =   765
         Value           =   1  'Checked
         Width           =   1095
      End
      Begin VB.CommandButton cmdBuscar 
         BackColor       =   &H00E0E0E0&
         Caption         =   "&Buscar"
         Default         =   -1  'True
         Height          =   915
         Left            =   11745
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   315
         Width           =   1050
      End
      Begin VB.CheckBox chkTodos 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Caption         =   "Todos"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   9660
         TabIndex        =   1
         Top             =   360
         Value           =   1  'Checked
         Width           =   1095
      End
      Begin MSComCtl2.DTPicker fdesde 
         Height          =   330
         Left            =   1575
         TabIndex        =   3
         Top             =   720
         Visible         =   0   'False
         Width           =   1545
         _ExtentX        =   2725
         _ExtentY        =   582
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
         Format          =   61210625
         CurrentDate     =   38002
      End
      Begin MSComCtl2.DTPicker fhasta 
         Height          =   330
         Left            =   4815
         TabIndex        =   4
         Top             =   720
         Visible         =   0   'False
         Width           =   1545
         _ExtentX        =   2725
         _ExtentY        =   582
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
         Format          =   61210625
         CurrentDate     =   38002
      End
      Begin pryCombo.miCombo cmbClientes 
         Height          =   330
         Left            =   1575
         TabIndex        =   18
         Top             =   315
         Width           =   7935
         _ExtentX        =   13996
         _ExtentY        =   582
      End
      Begin pryCombo.miCombo cmbMuestras 
         Height          =   330
         Left            =   1575
         TabIndex        =   19
         Top             =   720
         Width           =   7935
         _ExtentX        =   13996
         _ExtentY        =   582
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Tipo de Muestra"
         Height          =   195
         Index           =   3
         Left            =   225
         TabIndex        =   9
         Top             =   795
         Width           =   1185
      End
      Begin VB.Label lblCampos 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Prevista hasta"
         Height          =   195
         Index           =   2
         Left            =   3645
         TabIndex        =   7
         Top             =   780
         Visible         =   0   'False
         Width           =   1035
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Prevista desde"
         Height          =   195
         Index           =   1
         Left            =   225
         TabIndex        =   6
         Top             =   780
         Visible         =   0   'False
         Width           =   1065
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Cliente"
         Height          =   195
         Index           =   0
         Left            =   225
         TabIndex        =   5
         Top             =   390
         Width           =   495
      End
   End
   Begin MSComctlLib.ListView lista 
      Height          =   5475
      Left            =   60
      TabIndex        =   11
      Top             =   2565
      Width           =   13065
      _ExtentX        =   23045
      _ExtentY        =   9657
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
   Begin VB.Label lbltitulo 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "Listado de Muestras no cerradas del responsable : "
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
      Top             =   75
      Width           =   5355
   End
   Begin VB.Label lblsubtitulo 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "Listado de Muestras asignadas al responsable por fecha de entrega"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   90
      TabIndex        =   16
      Top             =   375
      Width           =   4785
   End
   Begin VB.Label lblCampos 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      Caption         =   "Pulse sobre el análisis para ver el detalle"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   210
      Index           =   8
      Left            =   4725
      TabIndex        =   12
      Top             =   8100
      Width           =   4095
   End
   Begin VB.Label lblmsg 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Caption         =   "Seleccione Criterio."
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   90
      TabIndex        =   8
      Top             =   2250
      Width           =   13065
   End
   Begin VB.Shape fondo 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   780
      Left            =   0
      Top             =   0
      Width           =   13155
   End
End
Attribute VB_Name = "frmTrabajo_Pendiente"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Const PK_ID_MUESTRA = 7
Private mvarCRITERIO_LISTADO As String

Private Sub cmdetiqueta_Click()
    ReDim ETIQUETAS(1)
    If lista.ListItems.Count = 0 Then Exit Sub
    ETIQUETAS(1) = lista.ListItems(lista.selectedItem.Index).SubItems(PK_ID_MUESTRA)
    frmEtiquetas.Show 1
End Sub


Private Sub cmdDeter_Click()
    If lista.ListItems.Count > 0 Then
        gmuestra = lista.ListItems(lista.selectedItem.Index).SubItems(PK_ID_MUESTRA)
        abrirRegistroMuestra gmuestra
'        Dim oMuestra As New clsMuestra
'        oMuestra.CargaMuestra (gmuestra)
'        Select Case oMuestra.getANALISIS_MODIFICADO
'            Case 2 ' Control de eficacia
'                With frmCE_Resultados
'                    .PK_ID_MUESTRA = gmuestra
'                    .Show 1
'                End With
'            Case 3 ' Sellante
'                frmSE_Resultados.Show 1
'            Case 5 ' Plasma
'                If oMuestra.getTIPO_MUESTRA_ID = TIPOS_MUESTRAS.DUREZA_ROCKWELL Or _
'                   oMuestra.getTIPO_MUESTRA_ID = TIPOS_MUESTRAS.DUREZA_BRINELL Or _
'                   oMuestra.getTIPO_MUESTRA_ID = TIPOS_MUESTRAS.DUREZA_VICKERS Then
'                    With frmPlasma_Dureza
'                        .PK = gmuestra
'                        .Show 1
'                    End With
'                Else
'                    With frmPlasma_Resultados
'                        .PK = gmuestra
'                        .Show 1
'                    End With
'                End If
'            Case Else
'                frmDeterminaciones.Show 1
'        End Select
        gmuestra = 0
    End If
End Sub

Private Sub cmdListado_Click()
    Dim objMuestras As clsMuestra
    
    If lista.ListItems.Count = 0 Then
        MsgBox "No existen registros para generar el listado.", vbExclamation, App.Title
        Set objMuestras = Nothing
        Exit Sub
    End If
    
    Set objMuestras = New clsMuestra
    
    objMuestras.imprimir_listadomuestras mvarCRITERIO_LISTADO, Format(fdesde.Value, "dd/mm/yyyy"), Format(fhasta.Value, "dd/mm/yyyy"), chkTodos.Value = vbChecked, chkTodas.Value = vbChecked, "Listado de Muestras Por Entregar"
    
    Set objMuestras = Nothing
End Sub

Private Sub chkTodas_Click()
    If chkTodas.Value = Checked Then
        cmbMuestras.limpiar
        cmbMuestras.desactivar
    Else
        cmbMuestras.activar
    End If
End Sub

Private Sub chkTodos_Click()
    If chkTodos.Value = Checked Then
        cmbClientes.limpiar
        cmbClientes.desactivar
    Else
        cmbClientes.activar
    End If
End Sub
Private Sub cmdSalir_Click()
    Unload Me
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = 27 Then
        cmdSalir_Click
    End If
End Sub
Private Sub Form_Load()
    log (Me.Name)
    cargar_botones Me
    Me.Left = 20
    Me.top = 20
    cabecera
    cargar_clientes
    cargar_muestras
    fdesde = Date
    fhasta = Date
    cmdBuscar_Click
    lbltitulo = lbltitulo & USUARIO.getUSUARIO
End Sub
Private Sub cabecera()
    With lista.ColumnHeaders.Add(, , "Código", 1000, lvwColumnLeft)
        .Tag = "Código"
    End With
    With lista.ColumnHeaders.Add(, , "Cliente", 2500, lvwColumnLeft)
        .Tag = "Cliente"
    End With
    With lista.ColumnHeaders.Add(, , "Analisis", 2400, lvwColumnLeft)
        .Tag = "Analisis"
    End With
    With lista.ColumnHeaders.Add(, , "Ref.Cliente", 2500, lvwColumnLeft)
        .Tag = "Ref.Cliente"
    End With
    With lista.ColumnHeaders.Add(, , "Recepción", 1050, lvwColumnCenter)
        .Tag = "Recepción"
    End With
    With lista.ColumnHeaders.Add(, , "Entrega", 1050, lvwColumnCenter)
        .Tag = "Entrega"
    End With
    With lista.ColumnHeaders.Add(, , "Id", 1, lvwColumnCenter)
        .Tag = "Id"
    End With
    With lista.ColumnHeaders.Add(, , "General", 1, lvwColumnCenter)
        .Tag = "General"
    End With
    With lista.ColumnHeaders.Add(, , "Responsable", 2000, lvwColumnLeft)
        .Tag = "Responsable"
    End With
End Sub
Private Sub cargar_clientes()
    llenar_combo cmbClientes, New clsCliente, 0, frmClientes, ""
End Sub
Private Sub cargar_muestras()
    llenar_combo cmbMuestras, New clsTipos_muestra, 0, frmTM_Detalle, "ANULADO = 0"
End Sub
Private Sub cmdBuscar_Click()
    Call buscar
End Sub
Private Sub buscar()
    Dim i As Integer
    Dim consulta As String
    Dim strMuestra As String
    Dim strClientes As String
    Dim strTipo As String
    On Error GoTo fallo
    lista.ListItems.Clear
    Dim rs As New ADODB.Recordset
    Dim f_desde As String
    Dim f_hasta As String
    f_desde = Format(fdesde, "yyyy-mm-dd")
    f_hasta = Format(fhasta, "yyyy-mm-dd")
    
    mvarCRITERIO_LISTADO = " 1 = 1 "
    If chkMias.Value = Checked Then
        mvarCRITERIO_LISTADO = " AND {muestras.RESPONSABLE_ID} = " & USUARIO.getID_EMPLEADO
    End If
    ' Tipo de muestra
    strMuestra = ""
    If chkTodas.Value = Unchecked Then
        If cmbMuestras.getPK_SALIDA = 0 Then
            MsgBox "Debe seleccionar un tipo de muestras.", vbExclamation, App.Title
            Exit Sub
        End If
        strMuestra = " AND mu.tipo_muestra_id=" & cmbMuestras.getPK_SALIDA
        mvarCRITERIO_LISTADO = mvarCRITERIO_LISTADO & " AND {muestras.TIPO_MUESTRA_ID} = " & cmbMuestras.getPK_SALIDA
    Else
        mvarCRITERIO_LISTADO = mvarCRITERIO_LISTADO & " AND {muestras.TIPO_MUESTRA_ID} > 0"
    End If
    ' Clientes
    strClientes = ""
    If chkTodos.Value = Unchecked Then
        If cmbClientes.getPK_SALIDA = 0 Then
            MsgBox "Debe seleccionar un cliente.", vbExclamation, App.Title
            Exit Sub
        End If
        strClientes = " AND mu.cliente_id = " & cmbClientes.getPK_SALIDA
        mvarCRITERIO_LISTADO = mvarCRITERIO_LISTADO & " AND {muestras.CLIENTE_ID} = " & cmbClientes.getPK_SALIDA
    Else
        mvarCRITERIO_LISTADO = mvarCRITERIO_LISTADO & " AND {muestras.CLIENTE_ID} > 0"
    End If
    ' Tipo
    strTipo = ""
    If chkCerradas.Value = Unchecked Then
        strTipo = " AND (mu.cerrada is Null or mu.cerrada = 0)"
        mvarCRITERIO_LISTADO = mvarCRITERIO_LISTADO & " AND (ISNULL({muestras.CERRADA}) OR {muestras.CERRADA}=0) "
    End If
    ' Fechas
    Dim fecha_desde As String
    fecha_desde = " AND mu.fecha_prev_fin>='" & f_desde & "'"
    fecha_desde = ""
    Dim fecha_hasta As String
    fecha_hasta = " AND mu.fecha_prev_fin<='" & f_hasta & "'"
    fecha_hasta = ""
    
    'las fechas previstas de fin de muestra se obvian
    'mvarCRITERIO_LISTADO = mvarCRITERIO_LISTADO & " AND {muestras.FECHA_PREV_FIN} >= Date(" & Year(f_desde) & "," & Month(f_desde) & "," & Day(f_desde) & ") AND {muestras.FECHA_PREV_FIN} <= Date(" & Year(f_hasta) & "," & Month(f_hasta) & "," & Day(f_hasta) & ")"
    
    
    consulta = "SELECT cl.id_cliente, " & _
               "concat(tm.codigo,'-',CAST(mu.id_particular AS CHAR)), " & _
               "cl.nombre, " & _
               "mu.tipo_analisis_id, " & _
               "mu.referencia_cliente, " & _
               "mu.fecha_recepcion, " & _
               "mu.id_muestra, " & _
               "mu.fecha_prev_fin, " & _
               "ta.nombre, " & _
               "mu.id_general, " & _
               "concat(emp.nombre,' ',emp.apellidos) " & _
               "FROM clientes as cl, tipos_muestra as tm,tipos_analisis as ta, " & _
                     "muestras as mu, " & _
                     "usuarios as emp " & _
               "WHERE mu.cliente_id=cl.id_cliente AND " & _
                      "mu.tipo_muestra_id=tm.id_tipo_muestra AND " & _
                      "mu.tipo_analisis_id=ta.id_tipo_analisis AND " & _
                      "tm.responsable_id = emp.id_empleado " & _
                      IIf(chkMias.Value = Checked, " AND mu.responsable_id = " & USUARIO.getID_EMPLEADO, "") & _
                      fecha_desde & _
                      fecha_hasta & _
                      strMuestra & _
                      strClientes & _
                      strTipo & _
                      " order by mu.fecha_prev_fin asc,mu.cliente_id,mu.id_muestra asc"
    Me.MousePointer = 11
    Set rs = datos_bd(consulta)
    If rs.RecordCount >= 1 Then
        lista.ListItems.Clear
        i = 1
        While Not rs.EOF
            With lista.ListItems.Add(, , rs.Fields(1))
            .SubItems(1) = rs.Fields(2)
            .SubItems(2) = rs.Fields(8)
            .SubItems(3) = rs.Fields(4)
            If Not IsNull(rs.Fields(5)) Then
            .SubItems(4) = Format(rs.Fields(5), "dd/mm/yyyy")
            End If
            If Not IsNull(rs.Fields(7)) Then
            .SubItems(5) = Format(rs.Fields(7), "dd/mm/yyyy")
            End If
            If Not IsNull(rs.Fields(9)) Then
            .SubItems(6) = rs.Fields(9)
            End If
            If Not IsNull(rs.Fields(6)) Then
            .SubItems(7) = rs.Fields(6)
            End If
            If Not IsNull(rs.Fields(10)) Then
            .SubItems(8) = rs.Fields(10)
            End If
            End With
            lista.ListItems(i).Checked = True
            i = i + 1
            rs.MoveNext
        Wend
        lblMsg.Caption = "Muestras entre el " & Format(fdesde, "dd/mm/yyyy") & " y " & Format(fhasta, "dd/mm/yyyy") & " (TOTAL : " & rs.RecordCount & ")"
    Else
        lblMsg.Caption = "No existe ninguna muestra con esos criterios."
    End If
    Me.MousePointer = 0
    Set rs = Nothing
    Exit Sub
fallo:
    Me.MousePointer = 0
    MsgBox "Se ha producido un error al buscar las muestras.", vbCritical, Err.Description
End Sub

Private Sub lista_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
   If lista.ListItems.Count > 0 Then
     lista.SortKey = ColumnHeader.Index - 1
     If lista.SortOrder = 0 Then
        lista.SortOrder = 1
     Else
        lista.SortOrder = 0
     End If
     lista.Sorted = True
   End If
End Sub

Private Sub lista_DblClick()
    gmuestra = lista.ListItems(lista.selectedItem.Index).SubItems(7)
    frmVerMuestra.Show 1
    actualizar_lista
End Sub


Public Sub actualizar_lista()
    ' Por si se ha modificado la muestra
    Dim consulta As String
    Dim rs As New ADODB.Recordset
    consulta = "SELECT cl.id_cliente, " & _
               "concat(tm.codigo,'-',CAST(mu.id_particular AS CHAR)), " & _
               "cl.nombre, " & _
               "mu.tipo_analisis_id, " & _
               "mu.referencia_cliente, " & _
               "mu.fecha_recepcion, " & _
               "mu.id_muestra, " & _
               "mu.fecha_prev_fin, " & _
               "mu.precio, " & _
               "mu.id_general " & _
               "FROM clientes as cl, " & _
                     "tipos_muestra as tm, " & _
                     "muestras as mu " & _
               "WHERE mu.cliente_id=cl.id_cliente AND " & _
                      "mu.tipo_muestra_id=tm.id_tipo_muestra AND " & _
                      "mu.id_muestra = " & CLng(lista.ListItems(lista.selectedItem.Index).SubItems(7))
    Set rs = datos_bd(consulta)
    If rs.RecordCount >= 1 Then
        Dim oMuestra As New clsMuestra
        lista.ListItems(lista.selectedItem.Index).Text = rs.Fields(1)
        lista.ListItems(lista.selectedItem.Index).SubItems(1) = rs.Fields(2)
        Dim oAnalisis As New clsTipos_analisis
        lista.ListItems(lista.selectedItem.Index).SubItems(2) = oAnalisis.NombreAnalisis(rs.Fields(3))
        Set oAnalisis = Nothing
        lista.ListItems(lista.selectedItem.Index).SubItems(3) = rs.Fields(4)
        If Not IsNull(rs.Fields(5)) Then
        lista.ListItems(lista.selectedItem.Index).SubItems(4) = rs.Fields(5)
        End If
        If Not IsNull(rs.Fields(7)) Then
        lista.ListItems(lista.selectedItem.Index).SubItems(5) = Format(rs.Fields(7), "dd/mm/yyyy")
        End If
    End If
    Set rs = Nothing
End Sub
'Private Sub cmdListado_Click_old()
'    Dim total As Currency
'    Dim i As Integer
'    On Error GoTo fallo
'    If lista.ListItems.Count = 0 Then
'        MsgBox "No existen registros para generar el listado.", vbExclamation, App.Title
'        Exit Sub
'    End If
'    Dim rs As New ADODB.RecordSet
'    rs.Fields.Append "c1", adChar, 15, adFldUpdatable
'    rs.Fields.Append "c2", adChar, 15, adFldUpdatable
'    rs.Fields.Append "c3", adChar, 50, adFldUpdatable
'    rs.Fields.Append "c4", adChar, 50, adFldUpdatable
'    rs.Fields.Append "c5", adChar, 50, adFldUpdatable
'    rs.Open
'    total = 0
'    For i = 1 To lista.ListItems.Count
'        rs.AddNew
'        rs("c1") = Left(lista.ListItems(i).SubItems(6) & "(" & lista.ListItems(i).Text & ")", 15)
'        rs("c2") = Left(lista.ListItems(i).SubItems(4), 15)
'        rs("c3") = Left(lista.ListItems(i).SubItems(1), 50)
'        rs("c4") = Left(lista.ListItems(i).SubItems(2), 50)
'        rs("c5") = Left(lista.ListItems(i).SubItems(3), 50)
'        rs.Update
'    Next
'    ' Generar Listado
'    Dim Listado As New dataListadoMuestras
'    ' Cabecera
'    With Listado.Sections("cabecera")
'        .Controls("lbltitulo").Caption = "Listado de Muestras a entregar desde " & Format(fdesde, "dd/mm/yyyy") & " al " & Format(fhasta, "dd/mm/yyyy")
'        If chkTodos.value = Checked Then
'            .Controls("lblcliente").Caption = "Cliente : *** TODOS ***"
'        Else
'            .Controls("lblcliente").Caption = "Cliente : " & cmbClientes.getTEXTO
'        End If
'    End With
'    Set Listado.Sections("cabecera").Controls("logo").Picture = LoadPicture(ReadINI(App.Path + "\config.ini", "logo", "logo"))
'    'Detalle
'    With Listado.Sections("detalle")
'        .Controls("c1").DataField = rs.Fields("c1").Name
'        .Controls("c2").DataField = rs.Fields("c2").Name
'        .Controls("c3").DataField = rs.Fields("c3").Name
'        .Controls("c4").DataField = rs.Fields("c4").Name
'        .Controls("c5").DataField = rs.Fields("c5").Name
'    End With
'    ' Pie de Pagina
''    With Listado.Sections("pie")
''        .Controls("lbltotal").Caption = Format(total, "currency")
''    End With
'    Set Listado.DataSource = rs
'    Listado.Caption = "Listado de Muestras Pendientes"
'    Listado.WindowState = vbMaximized
'    Listado.Show
'    Set rs = Nothing
''    Me.Height = 7890
''    Me.Width = 12780
'    Exit Sub
'fallo:
'    MsgBox "Error al generar el listado de Analisis pendientes.", vbCritical, Err.Description
'End Sub
