VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{CEA3B7F6-2847-4E5E-A551-DB7A62489D44}#38.0#0"; "miCombo.ocx"
Begin VB.Form frmTarifas 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Gestión de Tarifas"
   ClientHeight    =   9105
   ClientLeft      =   45
   ClientTop       =   300
   ClientWidth     =   12300
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "frmTarifas.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   9105
   ScaleWidth      =   12300
   Begin VB.CommandButton cmdImprimir 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Imprimir Tarifa de Precios"
      Height          =   870
      Left            =   5580
      Style           =   1  'Graphical
      TabIndex        =   25
      Top             =   8190
      Visible         =   0   'False
      Width           =   1950
   End
   Begin VB.CommandButton cmdexcel 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Generar Excel de Tarifas"
      Height          =   870
      Left            =   9000
      Picture         =   "frmTarifas.frx":000C
      Style           =   1  'Graphical
      TabIndex        =   24
      Top             =   8190
      Width           =   2085
   End
   Begin VB.TextBox txtdato 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   330
      Index           =   3
      Left            =   10080
      TabIndex        =   13
      Top             =   7785
      Width           =   1400
   End
   Begin VB.CommandButton cmdRecalculo 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Recalcular precio de muestras sin facturar"
      Height          =   870
      Left            =   2205
      Picture         =   "frmTarifas.frx":08D6
      Style           =   1  'Graphical
      TabIndex        =   18
      Top             =   8190
      Width           =   1725
   End
   Begin VB.CommandButton cmdDetalle 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Detalle"
      Height          =   870
      Left            =   45
      Picture         =   "frmTarifas.frx":11A0
      Style           =   1  'Graphical
      TabIndex        =   17
      Top             =   8190
      Width           =   1050
   End
   Begin VB.CommandButton cmdQuien 
      BackColor       =   &H00E0E0E0&
      Caption         =   "¿Donde?"
      Enabled         =   0   'False
      Height          =   870
      Left            =   1125
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   8190
      Width           =   1050
   End
   Begin VB.TextBox txtdato 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   330
      Index           =   2
      Left            =   8685
      Locked          =   -1  'True
      TabIndex        =   12
      Top             =   7785
      Width           =   1400
   End
   Begin VB.TextBox txtdato 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Height          =   330
      Index           =   1
      Left            =   7290
      Locked          =   -1  'True
      TabIndex        =   11
      Top             =   7785
      Width           =   1400
   End
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      ForeColor       =   &H80000008&
      Height          =   1725
      Left            =   45
      TabIndex        =   5
      Top             =   675
      Width           =   12165
      Begin VB.CommandButton cmdLimpiar 
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
         Height          =   645
         Left            =   11295
         Style           =   1  'Graphical
         TabIndex        =   21
         Top             =   675
         Width           =   735
      End
      Begin VB.OptionButton tipo 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Determinaciones"
         Height          =   195
         Index           =   2
         Left            =   9180
         TabIndex        =   9
         Top             =   270
         Width           =   1500
      End
      Begin VB.OptionButton tipo 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Baños"
         Height          =   195
         Index           =   1
         Left            =   8235
         TabIndex        =   8
         Top             =   270
         Width           =   960
      End
      Begin VB.OptionButton tipo 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Análisis"
         Height          =   195
         Index           =   0
         Left            =   7065
         TabIndex        =   7
         Top             =   270
         Width           =   960
      End
      Begin pryCombo.miCombo cmbtarifa 
         Height          =   330
         Left            =   810
         TabIndex        =   6
         Top             =   225
         Width           =   6135
         _ExtentX        =   10821
         _ExtentY        =   582
      End
      Begin pryCombo.miCombo cmbClientes 
         Height          =   330
         Left            =   810
         TabIndex        =   14
         Top             =   945
         Width           =   10275
         _ExtentX        =   18124
         _ExtentY        =   582
      End
      Begin pryCombo.miCombo cmbSolucion 
         Height          =   330
         Left            =   810
         TabIndex        =   20
         Top             =   1305
         Width           =   10275
         _ExtentX        =   18124
         _ExtentY        =   582
      End
      Begin pryCombo.miCombo cmbCodigo 
         Height          =   375
         Left            =   810
         TabIndex        =   22
         Top             =   585
         Width           =   10260
         _ExtentX        =   18098
         _ExtentY        =   661
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Código"
         Height          =   195
         Index           =   1
         Left            =   90
         TabIndex        =   23
         Top             =   630
         Width           =   495
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Solución"
         Height          =   195
         Left            =   90
         TabIndex        =   19
         Top             =   1395
         Width           =   615
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Cliente"
         Height          =   195
         Index           =   0
         Left            =   90
         TabIndex        =   15
         Top             =   990
         Width           =   480
      End
      Begin VB.Label Label1 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Tarifa"
         Height          =   285
         Left            =   90
         TabIndex        =   10
         Top             =   270
         Width           =   465
      End
   End
   Begin VB.TextBox txtdato 
      Appearance      =   0  'Flat
      Height          =   330
      Index           =   0
      Left            =   45
      Locked          =   -1  'True
      TabIndex        =   2
      Top             =   7785
      Width           =   7245
   End
   Begin VB.CommandButton cmdsalir 
      BackColor       =   &H00E0E0E0&
      Caption         =   "ESC-Salir"
      Height          =   870
      Left            =   11160
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   8190
      Width           =   1050
   End
   Begin MSComctlLib.ListView lista 
      Height          =   5325
      Left            =   45
      TabIndex        =   0
      Top             =   2430
      Width           =   12165
      _ExtentX        =   21458
      _ExtentY        =   9393
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
   Begin VB.Label lblsubtitulo 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "Listado y gestión de tarifas"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   90
      TabIndex        =   4
      Top             =   375
      Width           =   1875
   End
   Begin VB.Image imagen 
      Height          =   480
      Left            =   11475
      Picture         =   "frmTarifas.frx":1A6A
      Top             =   90
      Width           =   480
   End
   Begin VB.Label lbltitulo 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "Gestión de Tarifas"
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
      TabIndex        =   3
      Top             =   75
      Width           =   1935
   End
   Begin VB.Shape fondo 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   645
      Left            =   0
      Top             =   0
      Width           =   12315
   End
End
Attribute VB_Name = "frmTarifas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdImprimir_Click()
    With frmReport
         .iniciar
         .informe = "\Facturacion\rpttarifas_codigos"
'         If cmbfamilia.getTEXTO <> "" Then
'             .criterio = "{ado.FAMILIA_CODIGO_ID}=" & cmbfamilia.getPK_SALIDA
'         End If
         .imprimir = False
         .generar
         .Visible = True
    End With
End Sub
Private Sub cmbCodigo_change()
    cargar_lista
End Sub

Private Sub cmdExcel_Click()
    If cmbtarifa.getTEXTO = "" Then
        MsgBox "Seleccione una tarifa.", vbInformation, App.Title
        Exit Sub
    End If
     Dim rs_total As ADODB.RecordSet
     Dim XLA As Excel.Application
     Dim XLW As Excel.Workbook
     Dim XLS As Excel.Worksheet
   On Error GoTo cmdExcel_Click_Error

     Set XLA = New Excel.Application
     Set XLW = XLA.Workbooks.Add
     Set XLS = XLW.Worksheets(1)
     
     'Cabecera
     XLS.Cells(1, 1) = "Familia"
     XLS.Cells(1, 2) = "Código"
     XLS.Cells(1, 3) = "Descripción"
     XLS.Cells(1, 4) = "Cliente"
     XLS.Cells(1, 5) = cmbtarifa.getTEXTO
     
     XLS.Range("C1:C1").ColumnWidth = 60
     XLS.Range("D1:D1").ColumnWidth = 30
     XLS.Range("E1:E1").ColumnWidth = 14
     XLS.Range("A1:E1").Interior.ColorIndex = 35
         
     Dim oTarifas As New clsTarifas
     Dim rs As ADODB.RecordSet
'     Set rs = oTarifas.Listado()
'     Dim Col As Integer, fila As Integer
'     Col = 4
'     If rs.RecordCount > 0 Then
'        Do
'            XLS.Cells(1, Col) = rs(0)
'            Col = Col + 1
'            rs.MoveNext
'        Loop Until rs.EOF
'     End If
     ' Datos
     Dim oTarifas_Codigos As New clsTarifas_codigos
     Set rs = oTarifas_Codigos.Listado("", "", 0)
     Dim rs_banos As ADODB.RecordSet
     Dim aux As String
     Me.MousePointer = 11
     XLA.Visible = True
     If rs.RecordCount > 0 Then
        fila = 2
        Do
            XLS.Cells(fila, 1) = rs(3)
            XLS.Cells(fila, 2) = rs(1)
            XLS.Cells(fila, 3) = rs(2)
            XLS.Range("A" & CInt(fila) & ":E" & CInt(fila)).Interior.ColorIndex = 36
            fila = fila + 1
            ' Extraemos los componentes de cada código
            consulta = "SELECT a.nombre, b.codigo, cli.nombre, d.nombre, c.precio " & _
                   " FROM banos a " & _
                   " inner join tarifas_codigos b on a.tarifa_codigo_id = b.id_codigo " & _
                   " inner join tarifas_precios c on a.id_bano = c.bano_id and c.tipo_determinacion_id = 0 and c.tipo_analisis_id = 0 " & _
                   " inner join tarifas d on c.tarifa_id = d.id_tarifa " & _
                   " inner join clientes cli on a.cliente_id = cli.id_cliente " & _
                   " Where a.TARIFA_CODIGO_ID = " & rs(0) & _
                   "   and c.tarifa_id = " & cmbtarifa.getPK_SALIDA & _
                   " order by cli.nombre,a.nombre"
'            aux = ""
            Col = 4
            Set rs_banos = datos_bd(consulta)
            If rs_banos.RecordCount > 0 Then
 '               aux = rs_banos(0)
                Do
 '                   If aux <> rs_banos(0) Then
 '                       fila = fila + 1
 '                       Col = 4
 '                   Else
 '                       Col = Col + 1
 '                   End If
                    XLS.Cells(fila, 2) = rs_banos(1)
                    XLS.Cells(fila, 3) = rs_banos(0)
                    XLS.Cells(fila, 4) = rs_banos(2)
                    XLS.Cells(fila, 5) = rs_banos(4)
 '                   aux = rs_banos(0)
                    fila = fila + 1
                    rs_banos.MoveNext
                Loop Until rs_banos.EOF
            End If
            rs.MoveNext
        Loop Until rs.EOF
     End If
     XLS.Range("1:1").AutoFilter
     Me.MousePointer = 0
   On Error GoTo 0
   Exit Sub

cmdExcel_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure cmdexcel_Click of Formulario frmTarifas_Codigos"
End Sub

Private Sub cmdLimpiar_Click()
    cmbCodigo.Limpiar
    cmbClientes.Limpiar
    cmbSolucion.Limpiar
    cargar_lista
End Sub
Private Sub cmbClientes_change()
    cargar_lista
End Sub

Private Sub cmbSolucion_change()
    cargar_lista
End Sub

Private Sub cmbtarifa_change()
    cargar_lista
End Sub

Private Sub cmdDetalle_Click()
    If lista.ListItems.Count = 0 Then
        Exit Sub
    End If
    If tipo(0).value = True Then
        frmTA_Detalle.PK = lista.ListItems(lista.SelectedItem.Index).Text
        frmTA_Detalle.Show 1
    ElseIf tipo(1).value = True Then
        frmBANO_Detalle.PK = lista.ListItems(lista.SelectedItem.Index).Text
        frmBANO_Detalle.Show 1
    ElseIf tipo(2).value = True Then
        frmTD_Detalle.PK = lista.ListItems(lista.SelectedItem.Index).Text
        frmTD_Detalle.Show 1
    End If
End Sub
Private Sub cmdQuien_Click()
    If lista.ListItems.Count > 0 Then
        frmTD_Donde.PK = lista.ListItems(lista.SelectedItem.Index).Text
        frmTD_Donde.Show
    End If
End Sub
Private Sub cmdRecalculo_Click()
    If MsgBox("¿Esta seguro de recalcular el precio de las muestras sin facturar?", vbQuestion + vbYesNo, App.Title) = vbYes Then
        Dim oMuestra As New clsMuestra
        Me.MousePointer = 11
        If oMuestra.recalcular_precios_muestras_sin_facturar Then
            Me.MousePointer = 0
            MsgBox "Se han recalculado los precios correctamente.", vbInformation, App.Title
        End If
        Me.MousePointer = 0
    End If
End Sub

Private Sub cmdsalir_Click()
    Unload Me
End Sub

Private Sub Form_Activate()
    Me.SetFocus
End Sub
Private Sub Form_Load()
    log (Me.Name)
    Me.Top = 100
    Me.Left = 100
    cargar_botones Me
    cargar_combos
    cmbClientes.desactivar
    With lista.ColumnHeaders
        .Add , , "ID", 1, lvwColumnLeft
        .Add , , "Tipo análisis", 7200, lvwColumnLeft
        .Add , , "Código", 1400, lvwColumnCenter
        .Add , , "Tarifa Origen", 1400, lvwColumnRight
        .Add , , "Tarifa", 1400, lvwColumnRight
    End With
End Sub

Public Sub cargar_lista()
    Dim rs As ADODB.RecordSet
   On Error GoTo cargar_lista_Error

    Me.MousePointer = 11
    If tipo(0).value = True Then
        Dim oTA As New clsTipos_analisis
        If cmbCodigo.getTEXTO = "" Then
            Set rs = oTA.Listado_Tarifa(cmbtarifa.getPK_SALIDA, cmbClientes.getPK_SALIDA, 0)
        Else
            Set rs = oTA.Listado_Tarifa(cmbtarifa.getPK_SALIDA, cmbClientes.getPK_SALIDA, cmbCodigo.getPK_SALIDA)
        End If
    ElseIf tipo(1).value = True Then
        Dim oBANO As New clsBanos
        If cmbCodigo.getTEXTO = "" Then
            Set rs = oBANO.Listado_Tarifa(cmbtarifa.getPK_SALIDA, cmbClientes.getPK_SALIDA, cmbSolucion.getPK_SALIDA, 0)
        Else
            Set rs = oBANO.Listado_Tarifa(cmbtarifa.getPK_SALIDA, cmbClientes.getPK_SALIDA, cmbSolucion.getPK_SALIDA, cmbCodigo.getPK_SALIDA)
        End If
    ElseIf tipo(2).value = True Then
        Dim oTD As New clsTipos_determinacion
        If cmbCodigo.getTEXTO = "" Then
            Set rs = oTD.Listado_Tarifa(cmbtarifa.getPK_SALIDA, cmbClientes.getPK_SALIDA, cmbSolucion.getPK_SALIDA, 0)
        Else
            Set rs = oTD.Listado_Tarifa(cmbtarifa.getPK_SALIDA, cmbClientes.getPK_SALIDA, cmbSolucion.getPK_SALIDA, cmbCodigo.getPK_SALIDA)
        End If
    Else
        Me.MousePointer = 0
        Exit Sub
    End If
    ' Columnas de las tarifas
    Dim oTarifa As New clsTarifas
    oTarifa.Carga cmbtarifa.getPK_SALIDA
    lista.ColumnHeaders(5).Text = oTarifa.getNOMBRE  ' Tarifa seleccionada
    oTarifa.Carga oTarifa.getTARIFA_ORIGEN_ID
    lista.ColumnHeaders(4).Text = oTarifa.getNOMBRE ' Tarifa origen
    Dim i As Integer
    For i = 0 To 3
        txtdato(i) = ""
    Next
    lista.ListItems.Clear
    If rs.RecordCount <> 0 Then
        Do
            With lista.ListItems.Add(, , rs(0))
             .SubItems(1) = rs(1)
             If Not IsNull(rs(2)) Then
                .SubItems(2) = rs(2)
             End If
             .SubItems(3) = Format(rs(3), "currency")
             .SubItems(4) = Format(rs(4), "currency")
            End With
            rs.MoveNext
        Loop Until rs.EOF
        lista_Click
    End If
    Set oTP = Nothing
    Me.MousePointer = 0

   On Error GoTo 0
   Exit Sub

cargar_lista_Error:
    Me.MousePointer = 0

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure cargar_lista of Formulario frmTarifas"
End Sub

Private Sub lista_Click()
    If lista.ListItems.Count <> 0 Then
        txtdato(0) = lista.ListItems(lista.SelectedItem.Index).SubItems(1)
        txtdato(1) = lista.ListItems(lista.SelectedItem.Index).SubItems(2)
        txtdato(2) = lista.ListItems(lista.SelectedItem.Index).SubItems(3)
        txtdato(3) = lista.ListItems(lista.SelectedItem.Index).SubItems(4)
        On Error Resume Next
        txtdato(3).SetFocus
    End If
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
    cmdDetalle_Click
End Sub

Private Sub lista_KeyUp(KeyCode As Integer, Shift As Integer)
lista_Click
End Sub

Private Sub tipo_Click(Index As Integer)
    If cmbtarifa.getPK_SALIDA = 0 Then
        cmbtarifa.MostrarElemento 0
    End If
    cmbClientes.Limpiar
    cmdQuien.Enabled = False
    Select Case Index
    Case 0
        lista.ColumnHeaders(2).Text = "Tipos de análisis"
        cmbClientes.activar
        cmbSolucion.desactivar
    Case 1
        lista.ColumnHeaders(2).Text = "Baños"
        cmbClientes.activar
        cmbSolucion.activar
    Case 2
        lista.ColumnHeaders(2).Text = "Tipos de determinaciones"
        cmbClientes.activar
        cmbSolucion.activar
        cmdQuien.Enabled = True
    End Select
    cargar_lista
End Sub

Private Sub txtdato_GotFocus(Index As Integer)
    txtdato(Index).SelStart = 0
    txtdato(Index).SelLength = Len(txtdato(Index))
    txtdato(Index).BackColor = &H80C0FF
End Sub

Private Sub txtdato_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = 46 Then
       KeyAscii = 44
    End If
    If KeyAscii = 13 Then
        If Index = 3 Then
            anadir_precio
        Else
            txtdato(3).SetFocus
        End If
    End If
End Sub

Private Sub txtdato_LostFocus(Index As Integer)
    txtdato(Index).BackColor = vbWhite
    If Index = 2 Then
        txtdato(2) = Format(txtdato(2), "currency")
    End If
End Sub
Private Sub anadir_precio()
    If lista.ListItems.Count > 0 Then
        If txtdato(3) = "" Then
            MsgBox "Introduzca el precio correctamente.", vbInformation, App.Title
            txtdato(3).SetFocus
            Exit Sub
        End If
        If Not IsNumeric(3) Then
            MsgBox "Introduzca el precio correctamente.", vbInformation, App.Title
            txtdato(3).SetFocus
            Exit Sub
        End If
        lista.ListItems(lista.SelectedItem.Index).SubItems(4) = Format(txtdato(3), "currency")
        Dim oTP As New clsTarifas_precios
        Dim DETERMINACION As Long
        Dim ANALISIS As Long
        Dim BANO As Long
        If tipo(0).value = True Then
            ANALISIS = lista.ListItems(lista.SelectedItem.Index).Text
        End If
        If tipo(1).value = True Then
            BANO = lista.ListItems(lista.SelectedItem.Index).Text
        End If
        If tipo(2).value = True Then
            DETERMINACION = lista.ListItems(lista.SelectedItem.Index).Text
'            Dim oTD As New clsTipos_determinacion
'            oTD.Modificar_Tarifa CInt(determinacion), txtdato(1)
'            lista.ListItems(lista.SelectedItem.Index).SubItems(2) = txtdato(1)
        End If
        oTP.Modificar DETERMINACION, ANALISIS, BANO, cmbtarifa.getPK_SALIDA, Replace(txtdato(3), ",", ".")
        Set oTP = Nothing
        txttarifa = ""
        If lista.ListItems.Count > lista.SelectedItem.Index Then
                Set lista.SelectedItem = lista.ListItems(lista.SelectedItem.Index + 1)
                lista.SetFocus
                lista_Click
        End If
    End If
End Sub


Public Sub cargar_combos()
    llenar_combo cmbCodigo, New clsTarifas_codigos, 0, Me, ""
    llenar_combo cmbClientes, New clsCliente, 0, frmClientes, ""
    llenar_combo cmbtarifa, New clsTarifas, 0, Me, ""
    llenar_combo cmbSolucion, New clsSoluciones, 0, Me, ""
End Sub
