VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmConsultaBano 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Listado de Muestras especiales (Baños)"
   ClientHeight    =   7215
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   13095
   Icon            =   "frmConsultaBano.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7215
   ScaleWidth      =   13095
   Begin VB.CommandButton cmdListado 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Generar Listado"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   825
      Left            =   90
      Picture         =   "frmConsultaBano.frx":08CA
      Style           =   1  'Graphical
      TabIndex        =   17
      Top             =   6345
      Width           =   2085
   End
   Begin VB.CommandButton cmdsalir 
      BackColor       =   &H00E0E0E0&
      Caption         =   "ESC-Salir"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   870
      Left            =   12000
      Picture         =   "frmConsultaBano.frx":1194
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   6345
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
      Height          =   1635
      Left            =   45
      TabIndex        =   0
      Top             =   420
      Width           =   12990
      Begin VB.CheckBox chkTodas 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Caption         =   "Todas"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   10065
         TabIndex        =   13
         Top             =   720
         Width           =   1095
      End
      Begin VB.CommandButton cmdBuscar 
         BackColor       =   &H00E0E0E0&
         Caption         =   "&Buscar"
         Default         =   -1  'True
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   960
         Left            =   11430
         Picture         =   "frmConsultaBano.frx":149E
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   270
         Width           =   1410
      End
      Begin VB.CheckBox chkTodos 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Caption         =   "Todos"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   10065
         TabIndex        =   1
         Top             =   360
         Width           =   1095
      End
      Begin MSDataListLib.DataCombo cmbClientes 
         Height          =   360
         Left            =   2220
         TabIndex        =   3
         Top             =   300
         Width           =   7725
         _ExtentX        =   13626
         _ExtentY        =   635
         _Version        =   393216
         Appearance      =   0
         Text            =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MSComCtl2.DTPicker fdesde 
         Height          =   330
         Left            =   2220
         TabIndex        =   4
         Top             =   1110
         Width           =   1545
         _ExtentX        =   2725
         _ExtentY        =   582
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CalendarTitleBackColor=   14737632
         Format          =   50069505
         CurrentDate     =   38002
      End
      Begin MSComCtl2.DTPicker fhasta 
         Height          =   330
         Left            =   4500
         TabIndex        =   5
         Top             =   1110
         Width           =   1545
         _ExtentX        =   2725
         _ExtentY        =   582
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CalendarTitleBackColor=   14737632
         Format          =   50069505
         CurrentDate     =   38002
      End
      Begin MSDataListLib.DataCombo cmbMuestras 
         Height          =   360
         Left            =   2220
         TabIndex        =   12
         Top             =   690
         Width           =   7725
         _ExtentX        =   13626
         _ExtentY        =   635
         _Version        =   393216
         Appearance      =   0
         Text            =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Label lblCampos 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Tipo de Muestra"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   3
         Left            =   90
         TabIndex        =   11
         Top             =   750
         Width           =   1545
      End
      Begin VB.Label lblCampos 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "hasta"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   2
         Left            =   3870
         TabIndex        =   8
         Top             =   1170
         Width           =   555
      End
      Begin VB.Label lblCampos 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Recepcionadas desde"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   1
         Left            =   90
         TabIndex        =   7
         Top             =   1170
         Width           =   2085
      End
      Begin VB.Label lblCampos 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Cliente"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   0
         Left            =   90
         TabIndex        =   6
         Top             =   390
         Width           =   675
      End
   End
   Begin MSComctlLib.ListView lista 
      Height          =   3870
      Left            =   60
      TabIndex        =   14
      Top             =   2430
      Width           =   12990
      _ExtentX        =   22913
      _ExtentY        =   6826
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
      Left            =   4530
      TabIndex        =   15
      Top             =   6330
      Width           =   4095
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Listado de Muestras Especiales (Baños)"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   4
      Left            =   90
      TabIndex        =   10
      Top             =   45
      Width           =   12960
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
      Left            =   60
      TabIndex        =   9
      Top             =   2100
      Width           =   12990
   End
End
Attribute VB_Name = "frmConsultaBano"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub chkTodas_Click()
    If chkTodas.Value = Checked Then
        cmbMuestras.Text = ""
        cmbMuestras.Enabled = False
    Else
        cmbMuestras.Enabled = True
    End If
End Sub

Private Sub chkTodos_Click()
    If chkTodos.Value = Checked Then
        cmbClientes.Text = ""
        cmbClientes.Enabled = False
    Else
        cmbClientes.Enabled = True
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
    Me.Left = 20
    Me.Top = 20
    cabecera
    cargar_clientes
    cargar_muestras
    fdesde = Date
    fhasta = Date
End Sub
Public Sub cabecera()
    With lista.ColumnHeaders.Add(, , "Código", 800, lvwColumnLeft)
        .Tag = "Código"
    End With
    With lista.ColumnHeaders.Add(, , "Cliente", 2000, lvwColumnLeft)
        .Tag = "Cliente"
    End With
    With lista.ColumnHeaders.Add(, , "Baño", 2000, lvwColumnLeft)
        .Tag = "Baño"
    End With
    With lista.ColumnHeaders.Add(, , "Analisis", 2000, lvwColumnLeft)
        .Tag = "Analisis"
    End With
    With lista.ColumnHeaders.Add(, , "Ref.Cliente", 2100, lvwColumnLeft)
        .Tag = "Ref.Cliente"
    End With
    With lista.ColumnHeaders.Add(, , "Rutinario", 1000, lvwColumnCenter)
        .Tag = "Rutinario"
    End With
    With lista.ColumnHeaders.Add(, , "Fecha", 1100, lvwColumnCenter)
        .Tag = "Fecha"
    End With
    With lista.ColumnHeaders.Add(, , "Precio", 1000, lvwColumnRight)
        .Tag = "Precio"
    End With
        With lista.ColumnHeaders.Add(, , "Id", 700, lvwColumnCenter)
        .Tag = "Id"
    End With
End Sub
Public Sub cargar_clientes()
    Dim oCliente As New clsCliente
    Set cmbClientes.RowSource = oCliente.Clientes_muestras_especiales  'recorset devuelto por la funcion
    cmbClientes.ListField = "nombre" 'campo que veo
    cmbClientes.DataField = "id_cliente" 'campo asociado
    cmbClientes.BoundColumn = "id_cliente" 'lo que realmente envia
    Set oCliente = Nothing
End Sub

Public Sub cargar_muestras()
    Dim omuestra As New clsTipos_muestra
    Set cmbMuestras.RowSource = omuestra.Listado_especiales
    cmbMuestras.ListField = "nombre" 'lo que enseña
    cmbMuestras.DataField = "id_tipo_muestra" 'campo asociado
    cmbMuestras.BoundColumn = "id_tipo_muestra" 'lo que realmente envia
    Set omuestra = Nothing
End Sub
Private Sub cmdBuscar_Click()
    Call buscar
End Sub
Private Sub buscar()
    Dim i As Integer
    Dim consulta As String
    Dim strMuestra As String
    Dim strClientes As String
    On Error GoTo fallo
    lista.ListItems.Clear
    Dim rs As New ADODB.Recordset
    Dim f_desde As String
    Dim f_hasta As String
    f_desde = Format(fdesde, "yyyy-mm-dd")
    f_hasta = Format(fhasta, "yyyy-mm-dd")
    ' Tipo de muestra
    strMuestra = ""
    If chkTodas.Value = Unchecked Then
        If cmbMuestras.Text = "" Then
            MsgBox "Debe seleccionar un tipo de muestras.", vbExclamation, App.Title
            Exit Sub
        End If
        strMuestra = " AND mu.tipo_muestra_id=" & cmbMuestras.BoundText
    End If
    ' Clientes
    strClientes = ""
    If chkTodos.Value = Unchecked Then
        If cmbClientes.Text = "" Then
            MsgBox "Debe seleccionar un cliente.", vbExclamation, App.Title
            Exit Sub
        End If
        strClientes = " AND mu.cliente_id = " & cmbClientes.BoundText
    End If
    ' Fechas
    Dim fecha_desde As String
    fecha_desde = " AND mu.fecha_recepcion>='" & f_desde & "'"
    Dim fecha_hasta As String
    fecha_hasta = " AND mu.fecha_recepcion<='" & f_hasta & "'"
    consulta = "SELECT cl.id_cliente, " & _
               "concat(tm.codigo,'-',mu.id_particular), " & _
               "cl.nombre, " & _
               "mu.tipo_analisis_id, " & _
               "mu.referencia_cliente, " & _
               "mu.fecha_recepcion, " & _
               "mu.id_muestra, " & _
               "mu.precio, " & _
               "b.nombre " & _
               "FROM clientes as cl, " & _
                     "tipos_muestra as tm, " & _
                     "muestras as mu, " & _
                     "banos as b " & _
               "WHERE mu.cliente_id=cl.id_cliente AND " & _
                      "mu.tipo_muestra_id=tm.id_tipo_muestra and " & _
                      "tm.tipo_especial_id <> 0 and " & _
                      "b.id_bano = mu.bano_id " & _
                      fecha_desde & _
                      fecha_hasta & _
                      strMuestra & _
                      strClientes & _
                      " order by b.nombre,mu.id_muestra"
    Me.MousePointer = 11
    Set rs = datos_bd(consulta)
    If rs.RecordCount >= 1 Then
        Dim omuestra As New clsMuestra
        Dim oAnalisis As New clsTipos_analisis
        Dim ote As New clsValores_banos
        lista.ListItems.Clear
        i = 1
        While Not rs.EOF
            With lista.ListItems.Add(, , rs.Fields(1))
            .SubItems(1) = rs.Fields(2)
            .SubItems(2) = rs.Fields(8)
            .SubItems(3) = oAnalisis.NombreAnalisis(rs.Fields(3))
            .SubItems(4) = rs.Fields(4)
            ' Rutinario / Recarga
            .SubItems(5) = ote.rutinario(rs.Fields(6))
            ' Fecha
            If Not IsNull(rs.Fields(5)) Then
            .SubItems(6) = rs.Fields(5)
            End If
            ' Precio
            If Not IsNull(rs.Fields(7)) Then
            .SubItems(7) = Format(rs.Fields(7), "currency")
            End If
            ' Id
            .SubItems(8) = rs.Fields(6)
            End With
            lista.ListItems(i).Checked = True
            i = i + 1
            rs.MoveNext
        Wend
        lblmsg.Caption = "Muestras entre el " & Format(fdesde, "dd/mm/yyyy") & " y " & Format(fhasta, "dd/mm/yyyy")
    Else
        lblmsg.Caption = "No existe ninguna muestra con esos criterios."
    End If
    Set oAnalisis = Nothing
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
    If lista.ListItems.Count > 0 Then
        gmuestra = lista.ListItems(lista.SelectedItem.Index).SubItems(8)
        frmVerMuestra.Show 1
'        actualizar_lista
    End If
End Sub

Public Sub actualizar_lista()
    ' Por si se ha modificado la muestra
    Dim consulta As String
    Dim rs As New ADODB.Recordset
    consulta = "SELECT cl.id_cliente, " & _
               "concat(tm.codigo,'-',mu.id_particular), " & _
               "cl.nombre, " & _
               "mu.tipo_analisis_id, " & _
               "mu.referencia_cliente, " & _
               "mu.fecha_recepcion, " & _
               "mu.id_muestra, " & _
               "mu.precio " & _
               "FROM clientes as cl, " & _
                     "tipos_muestra as tm, " & _
                     "muestras as mu " & _
               "WHERE mu.cliente_id=cl.id_cliente AND " & _
                      "mu.tipo_muestra_id=tm.id_tipo_muestra AND " & _
                      "mu.id_muestra = " & CLng(lista.ListItems(lista.SelectedItem.Index).SubItems(8))
    Set rs = datos_bd(consulta)
    If rs.RecordCount >= 1 Then
        Dim omuestra As New clsMuestra
        lista.ListItems(lista.SelectedItem.Index).Text = rs.Fields(1)
        lista.ListItems(lista.SelectedItem.Index).SubItems(1) = rs.Fields(2)
        Dim oAnalisis As New clsTipos_analisis
        lista.ListItems(lista.SelectedItem.Index).SubItems(2) = oAnalisis.NombreAnalisis(rs.Fields(3))
        Set oAnalisis = Nothing
        lista.ListItems(lista.SelectedItem.Index).SubItems(3) = rs.Fields(4)
        If Not IsNull(rs.Fields(5)) Then
        lista.ListItems(lista.SelectedItem.Index).SubItems(4) = rs.Fields(5)
        End If
        If Not IsNull(rs.Fields(7)) Then
        lista.ListItems(lista.SelectedItem.Index).SubItems(5) = Format(rs.Fields(7), "currency")
        End If
    End If
    Set rs = Nothing
End Sub
Private Sub cmdListado_Click()
    Dim total As Currency
    Dim total_general As Currency
    Dim i As Integer
    On Error GoTo fallo
    If lista.ListItems.Count = 0 Then
        MsgBox "No existen registros para generar el listado.", vbExclamation, App.Title
        Exit Sub
    End If
    Dim rs As New ADODB.Recordset
    rs.Fields.Append "c1", adChar, 30, adFldUpdatable
    rs.Fields.Append "c2", adChar, 50, adFldUpdatable
    rs.Fields.Append "c3", adChar, 50, adFldUpdatable
    rs.Fields.Append "c4", adChar, 10, adFldUpdatable
    rs.Fields.Append "c5", adChar, 15, adFldUpdatable
    rs.Open
    total = 0
    total_general = 0
    Dim bano_ant As String
    bano_ant = ""
    For i = 1 To lista.ListItems.Count
        rs.AddNew
        If bano_ant <> lista.ListItems(i).SubItems(2) Then
            If bano_ant <> "" Then
                rs("c4") = "SUBTOTAL"
                rs("c5") = Format(total, "currency")
                total = 0
                rs.Update
                rs.AddNew
            End If
            rs("c1") = Left(lista.ListItems(i).SubItems(2), 30)
            bano_ant = lista.ListItems(i).SubItems(2)
        End If
        rs("c2") = Left(lista.ListItems(i).SubItems(3), 50)
        rs("c3") = Left(lista.ListItems(i).SubItems(4), 50)
        rs("c4") = Left(lista.ListItems(i).SubItems(5), 10)
        rs("c5") = Left(lista.ListItems(i).SubItems(7), 15)
        total = total + CCur(rs("c5"))
        total_general = total_general + total
        rs.Update
    Next
    ' Generar Listado
    Dim Listado As New rptConsultaBano
    ' Cabecera
    With Listado.Sections("cabecera")
        .Controls("lbltitulo").Caption = "Listado de Análisis de Muestras Especiales desde " & Format(fdesde, "dd/mm/yyyy") & " al " & Format(fhasta, "dd/mm/yyyy")
        If chkTodos.Value = Checked Then
            .Controls("lblcliente").Caption = "Cliente : *** TODOS ***"
        Else
            .Controls("lblcliente").Caption = "Cliente : " & cmbClientes.Text
        End If
    End With
    'Detalle
    With Listado.Sections("detalle")
        .Controls("c1").DataField = rs.Fields("c1").Name
        .Controls("c2").DataField = rs.Fields("c2").Name
        .Controls("c3").DataField = rs.Fields("c3").Name
        .Controls("c4").DataField = rs.Fields("c4").Name
        .Controls("c5").DataField = rs.Fields("c5").Name
    End With
    ' Pie de Pagina
    With Listado.Sections("pie")
        .Controls("lbltotal").Caption = Format(total_general, "currency")
    End With
    Set Listado.DataSource = rs
    Listado.Caption = "Listado de Análisis de Muestras Especiales"
    Listado.WindowState = vbMaximized
    Listado.Show
    Set rs = Nothing
'    Me.Height = 7890
'    Me.Width = 12780
    Exit Sub
fallo:
    MsgBox "Error al generar el listado de Muestras Especiales.", vbCritical, Err.Description
End Sub
