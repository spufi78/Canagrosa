VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{CEA3B7F6-2847-4E5E-A551-DB7A62489D44}#38.0#0"; "miCombo.ocx"
Begin VB.Form frmTarifas_Codigos_Asignacion 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Gestión de Tarifas"
   ClientHeight    =   9795
   ClientLeft      =   45
   ClientTop       =   300
   ClientWidth     =   15030
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "frmTarifas_Codigos_Asignacion.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9795
   ScaleWidth      =   15030
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame2 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      ForeColor       =   &H80000008&
      Height          =   1725
      Left            =   7515
      TabIndex        =   15
      Top             =   765
      Width           =   7485
      Begin VB.OptionButton tipo 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Análisis"
         Height          =   195
         Index           =   0
         Left            =   5895
         TabIndex        =   19
         Top             =   270
         Width           =   960
      End
      Begin VB.OptionButton tipo 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Baños"
         Height          =   195
         Index           =   1
         Left            =   5895
         TabIndex        =   18
         Top             =   540
         Value           =   -1  'True
         Width           =   960
      End
      Begin VB.OptionButton tipo 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Determinaciones"
         Height          =   195
         Index           =   2
         Left            =   5895
         TabIndex        =   17
         Top             =   810
         Width           =   1500
      End
      Begin VB.CommandButton cmdLimpiar 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Limpiar Campos"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   5895
         Style           =   1  'Graphical
         TabIndex        =   16
         Top             =   1170
         Width           =   1455
      End
      Begin pryCombo.miCombo cmbtarifa 
         Height          =   330
         Left            =   810
         TabIndex        =   20
         Top             =   225
         Width           =   4965
         _ExtentX        =   8758
         _ExtentY        =   582
      End
      Begin pryCombo.miCombo cmbClientes 
         Height          =   330
         Left            =   810
         TabIndex        =   21
         Top             =   945
         Width           =   4965
         _ExtentX        =   8758
         _ExtentY        =   582
      End
      Begin pryCombo.miCombo cmbSolucion 
         Height          =   330
         Left            =   810
         TabIndex        =   22
         Top             =   1305
         Width           =   4965
         _ExtentX        =   8758
         _ExtentY        =   582
      End
      Begin pryCombo.miCombo cmbCodigo 
         Height          =   375
         Left            =   810
         TabIndex        =   23
         Top             =   585
         Width           =   4965
         _ExtentX        =   8758
         _ExtentY        =   661
      End
      Begin VB.Label Label4 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Tarifa"
         Height          =   285
         Left            =   90
         TabIndex        =   27
         Top             =   270
         Width           =   465
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Cliente"
         Height          =   195
         Index           =   0
         Left            =   90
         TabIndex        =   26
         Top             =   990
         Width           =   480
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Solución"
         Height          =   195
         Index           =   2
         Left            =   90
         TabIndex        =   25
         Top             =   1395
         Width           =   615
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Código"
         Height          =   195
         Index           =   1
         Left            =   90
         TabIndex        =   24
         Top             =   630
         Width           =   495
      End
   End
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   "Filtro"
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
      Height          =   1140
      Left            =   45
      TabIndex        =   4
      Top             =   765
      Width           =   7440
      Begin VB.TextBox txtDatos 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Index           =   1
         Left            =   4185
         MaxLength       =   255
         TabIndex        =   10
         Top             =   675
         Width           =   2250
      End
      Begin VB.TextBox txtDatos 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Index           =   0
         Left            =   855
         MaxLength       =   255
         TabIndex        =   8
         Top             =   675
         Width           =   1215
      End
      Begin pryCombo.miCombo cmbfamilia 
         Height          =   330
         Left            =   855
         TabIndex        =   5
         Top             =   270
         Width           =   6315
         _ExtentX        =   11139
         _ExtentY        =   582
      End
      Begin VB.Label Label2 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Descripción"
         Height          =   285
         Index           =   1
         Left            =   3195
         TabIndex        =   9
         Top             =   720
         Width           =   915
      End
      Begin VB.Label Label2 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Código"
         Height          =   285
         Index           =   0
         Left            =   90
         TabIndex        =   7
         Top             =   720
         Width           =   690
      End
      Begin VB.Label Label1 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Familia"
         Height          =   285
         Left            =   90
         TabIndex        =   6
         Top             =   315
         Width           =   690
      End
   End
   Begin VB.CommandButton cmdsalir 
      BackColor       =   &H00E0E0E0&
      Caption         =   "ESC-Salir"
      Height          =   870
      Left            =   13905
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   8865
      Width           =   1050
   End
   Begin MSComctlLib.ListView lista 
      Height          =   6810
      Left            =   45
      TabIndex        =   1
      Top             =   1980
      Width           =   7440
      _ExtentX        =   13123
      _ExtentY        =   12012
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
   Begin MSComctlLib.ListView noasignados 
      Height          =   2850
      Left            =   7515
      TabIndex        =   11
      Top             =   2790
      Width           =   7440
      _ExtentX        =   13123
      _ExtentY        =   5027
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      AllowReorder    =   -1  'True
      Checkboxes      =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   14737632
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
   Begin MSComctlLib.ListView asignados 
      Height          =   2850
      Left            =   7515
      TabIndex        =   12
      Top             =   5940
      Width           =   7440
      _ExtentX        =   13123
      _ExtentY        =   5027
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      AllowReorder    =   -1  'True
      Checkboxes      =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   14737632
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
   Begin VB.Label Label2 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Listado de baños asignados a ese código."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   285
      Index           =   5
      Left            =   7515
      TabIndex        =   14
      Top             =   5670
      Width           =   7440
   End
   Begin VB.Label Label2 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Listado de baños sin código asignado."
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
      Index           =   4
      Left            =   7515
      TabIndex        =   13
      Top             =   2520
      Width           =   7440
   End
   Begin VB.Label lblsubtitulo 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "Seleccione un código tarifario y asigneselo a los análisis correspondientes"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   90
      TabIndex        =   3
      Top             =   420
      Width           =   5205
   End
   Begin VB.Image imagen 
      Height          =   480
      Left            =   14355
      Picture         =   "frmTarifas_Codigos_Asignacion.frx":000C
      Top             =   90
      Width           =   480
   End
   Begin VB.Label lbltitulo 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "Asignación de códigos tarifarios a Tipos de análisis, baños y determinaciones"
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
      TabIndex        =   2
      Top             =   120
      Width           =   8160
   End
   Begin VB.Shape fondo 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   735
      Left            =   0
      Top             =   0
      Width           =   14970
   End
End
Attribute VB_Name = "frmTarifas_Codigos_Asignacion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub asignados_DblClick()
    If asignados.ListItems.Count = 0 Then
        Exit Sub
    End If
    If tipo(0).value = True Then
        frmTA_Detalle.PK = asignados.ListItems(asignados.SelectedItem.Index).SubItems(3)
        frmTA_Detalle.Show 1
    ElseIf tipo(1).value = True Then
        frmBANO_Detalle.PK = asignados.ListItems(asignados.SelectedItem.Index).SubItems(3)
        frmBANO_Detalle.Show 1
    ElseIf tipo(2).value = True Then
        frmTD_Detalle.PK = asignados.ListItems(asignados.SelectedItem.Index).SubItems(3)
        frmTD_Detalle.Show 1
    End If

End Sub

Private Sub asignados_ItemCheck(ByVal Item As MSComctlLib.ListItem)
    If asignados.ListItems.Count > 0 Then
        Dim i As Integer
        Dim seleccionado As Integer
        For i = 1 To asignados.ListItems.Count
            If asignados.ListItems(i).Checked = True Then
                seleccionado = i
            End If
        Next
        If tipo(0).value = True Then
            Dim oTA As New clsTipos_analisis
            oTA.setTARIFA_CODIGO_ID = 0
            oTA.Modificar_Codigo_Tarifa asignados.ListItems(seleccionado).SubItems(3)
        End If
        If tipo(1).value = True Then
            Dim oBANO As New clsBanos
            oBANO.setTARIFA_CODIGO_ID = 0
            oBANO.Modificar_Codigo_Tarifa asignados.ListItems(seleccionado).SubItems(3)
        End If
        If tipo(2).value = True Then
            Dim oTD As New clsTipos_determinacion
            oTD.setTARIFA_CODIGO_ID = 0
            oTD.Modificar_Codigo_Tarifa asignados.ListItems(seleccionado).SubItems(3)
        End If
        cargar_relacionados
    End If
End Sub

Private Sub cmbClientes_change()
    cargar_relacionados
End Sub

Private Sub cmbCodigo_change()
    cargar_relacionados
End Sub

Private Sub cmbfamilia_Change()
    cargar_lista
End Sub

Private Sub cmbSolucion_change()
    cargar_relacionados
End Sub

Private Sub cmbtarifa_change()
    cargar_relacionados
End Sub

Private Sub cmdLimpiar_Click()
    cmbCodigo.Limpiar
    cmbClientes.Limpiar
    cmbSolucion.Limpiar
    cargar_relacionados
End Sub

Private Sub cmdSalir_Click()
    Unload Me
End Sub

Private Sub Form_Activate()
    Me.SetFocus
'    cargar_lista
End Sub
Private Sub Form_Load()
    log (Me.Name)
    cargar_botones Me
    cargar_combos
    With lista.ColumnHeaders
        .Add , , "ID", 1, lvwColumnLeft
        .Add , , "Código", 1000, lvwColumnCenter
        .Add , , "Descripción", 4500, lvwColumnLeft
        .Add , , "Familia", 1200, lvwColumnCenter
        .Add , , "id_familia", 1, lvwColumnRight
    End With
    With asignados.ColumnHeaders
        .Add , , "Baño", 4500, lvwColumnLeft
        .Add , , "Código", 1200, lvwColumnCenter
        .Add , , "Tarifa", 1200, lvwColumnRight
        .Add , , "ID", 1, lvwColumnLeft
    End With
    With noasignados.ColumnHeaders
        .Add , , "Baño", 4500, lvwColumnLeft
        .Add , , "Código", 1200, lvwColumnCenter
        .Add , , "Tarifa", 1200, lvwColumnRight
        .Add , , "ID", 1, lvwColumnLeft
    End With
    cargar_lista
End Sub

Public Sub cargar_lista()
    Dim rs As ADODB.RecordSet
    Dim oTC As New clsTarifas_codigos
    Set rs = oTC.Listado_Filtro(cmbFamilia.getPK_SALIDA, txtdatos(0), txtdatos(1))
    lista.ListItems.Clear
    If rs.RecordCount <> 0 Then
        Do
            With lista.ListItems.Add(, , rs(0))
             .SubItems(1) = rs(1)
             .SubItems(2) = rs(2)
             .SubItems(3) = rs(3)
             .SubItems(4) = rs(4)
            End With
            rs.MoveNext
        Loop Until rs.EOF
        lista_Click
    End If
    Set oTC = Nothing
End Sub

Private Sub lista_Click()
    cargar_relacionados
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
Public Sub cargar_combos()
    llenar_combo cmbFamilia, New clsTarifas_codigos_familias, 0, Me, ""
    llenar_combo cmbCodigo, New clsTarifas_codigos, 0, Me, ""
    llenar_combo cmbClientes, New clsCliente, 0, frmClientes, ""
    llenar_combo cmbTarifa, New clsTarifas, 0, Me, ""
    llenar_combo cmbSolucion, New clsSoluciones, 0, Me, ""
    cmbTarifa.MostrarElemento 0
End Sub

Private Sub lista_KeyUp(KeyCode As Integer, Shift As Integer)
    lista_Click
End Sub

Private Sub noasignados_DblClick()
    If noasignados.ListItems.Count = 0 Then
        Exit Sub
    End If
    If tipo(0).value = True Then
        frmTA_Detalle.PK = noasignados.ListItems(noasignados.SelectedItem.Index).SubItems(3)
        frmTA_Detalle.Show 1
    ElseIf tipo(1).value = True Then
        frmBANO_Detalle.PK = noasignados.ListItems(noasignados.SelectedItem.Index).SubItems(3)
        frmBANO_Detalle.Show 1
    ElseIf tipo(2).value = True Then
        frmTD_Detalle.PK = noasignados.ListItems(noasignados.SelectedItem.Index).SubItems(3)
        frmTD_Detalle.Show 1
    End If

End Sub

Private Sub noasignados_ItemCheck(ByVal Item As MSComctlLib.ListItem)
    If noasignados.ListItems.Count > 0 And lista.ListItems.Count > 0 Then
        Dim i As Integer
        Dim seleccionado As Integer
        For i = 1 To noasignados.ListItems.Count
            If noasignados.ListItems(i).Checked = True Then
'                MsgBox i
                seleccionado = i
            End If
        Next
        If tipo(0).value = True Then
            Dim oTA As New clsTipos_analisis
            oTA.setTARIFA_CODIGO_ID = lista.ListItems(lista.SelectedItem.Index)
            oTA.Modificar_Codigo_Tarifa noasignados.ListItems(seleccionado).SubItems(3)
        End If
        If tipo(1).value = True Then
            Dim oBANO As New clsBanos
            oBANO.setTARIFA_CODIGO_ID = lista.ListItems(lista.SelectedItem.Index)
            oBANO.Modificar_Codigo_Tarifa noasignados.ListItems(seleccionado).SubItems(3)
        End If
        If tipo(2).value = True Then
            Dim oTD As New clsTipos_determinacion
            oTD.setTARIFA_CODIGO_ID = lista.ListItems(lista.SelectedItem.Index)
            oTD.Modificar_Codigo_Tarifa noasignados.ListItems(seleccionado).SubItems(3)
        End If
        cargar_relacionados
    End If
End Sub

Private Sub tipo_Click(Index As Integer)
    If cmbTarifa.getPK_SALIDA = 0 Then
        cmbTarifa.MostrarElemento 0
    End If
    cmbClientes.Limpiar
    Select Case Index
    Case 0
        Label2(4) = "Listado de T.A. sin código asignado."
        Label2(5) = "Listado de T.A. asignados a ese código."
        asignados.ColumnHeaders(1).Text = "Tipos de análisis"
        noasignados.ColumnHeaders(1).Text = "Tipos de análisis"
        cmbClientes.activar
        cmbSolucion.desactivar
    Case 1
        Label2(4) = "Listado de baños sin código asignado."
        Label2(5) = "Listado de baños asignados a ese código."
        asignados.ColumnHeaders(1).Text = "Baños"
        noasignados.ColumnHeaders(1).Text = "Baños"
        cmbClientes.activar
        cmbSolucion.activar
    Case 2
        Label2(4) = "Listado de determinaciones sin código asignado."
        Label2(5) = "Listado de determinaciones asignadas a ese código."
        asignados.ColumnHeaders(1).Text = "Tipos de determinaciones"
        noasignados.ColumnHeaders(1).Text = "Tipos de determinaciones"
        cmbClientes.activar
        cmbSolucion.activar
    End Select
    cargar_relacionados
End Sub

Private Sub txtDatos_Change(Index As Integer)
    cargar_lista
End Sub

Private Sub cargar_relacionados()
    On Error Resume Next
    Dim rs As ADODB.RecordSet
    Dim codigo As Long
    Dim cliente As Long
    Dim solucion As Long
    If cmbCodigo.getTEXTO = "" Then
        codigo = 0
    Else
        codigo = cmbCodigo.getPK_SALIDA
    End If
    If cmbClientes.getTEXTO = "" Then
        cliente = 0
    Else
        cliente = cmbClientes.getPK_SALIDA
    End If
    If cmbSolucion.getTEXTO = "" Then
        solucion = 0
    Else
        solucion = cmbSolucion.getPK_SALIDA
    End If
    If tipo(0).value = True Then
        Dim oTA As New clsTipos_analisis
        Set rs = oTA.Listado_Tarifa(cmbTarifa.getPK_SALIDA, cliente, codigo)
    ElseIf tipo(1).value = True Then
        Dim oBANO As New clsBanos
        Set rs = oBANO.Listado_Tarifa(cmbTarifa.getPK_SALIDA, cliente, solucion, codigo)
    ElseIf tipo(2).value = True Then
        Dim oTD As New clsTipos_determinacion
        Set rs = oTD.Listado_Tarifa(cmbTarifa.getPK_SALIDA, cliente, solucion, codigo)
    End If
    asignados.ColumnHeaders(3).Text = cmbTarifa.getTEXTO  ' Tarifa seleccionada
    noasignados.ColumnHeaders(3).Text = cmbTarifa.getTEXTO  ' Tarifa seleccionada
    asignados.ListItems.Clear
    noasignados.ListItems.Clear
    If rs.RecordCount <> 0 Then
        Do
          If CInt(rs(5)) = CInt(lista.ListItems(lista.SelectedItem.Index)) Then
            With asignados.ListItems.Add(, , rs(1))
             If Not IsNull(rs(2)) Then
                .SubItems(1) = rs(2)
             End If
             .SubItems(2) = Format(rs(4), "currency")
             .SubItems(3) = rs(0)
            End With
          Else
           If CInt(rs(5)) = 0 Then
            With noasignados.ListItems.Add(, , rs(1))
             If Not IsNull(rs(2)) Then
                .SubItems(1) = rs(2)
             End If
             .SubItems(2) = Format(rs(4), "currency")
             .SubItems(3) = rs(0)
            End With
           End If
          End If
          rs.MoveNext
        Loop Until rs.EOF
    End If
    Set oTP = Nothing
End Sub
