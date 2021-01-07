VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{CEA3B7F6-2847-4E5E-A551-DB7A62489D44}#46.0#0"; "miCombo.ocx"
Begin VB.Form frmInventario_Listado 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Inventario"
   ClientHeight    =   9990
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   15255
   Icon            =   "frmInventario_Listado.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   9990
   ScaleWidth      =   15255
   Begin VB.CommandButton cmdExcel 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Excel"
      Height          =   870
      Left            =   6345
      Style           =   1  'Graphical
      TabIndex        =   22
      Top             =   9045
      Width           =   1230
   End
   Begin VB.CommandButton cmdDuplicar 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Duplicar"
      Height          =   870
      Left            =   5085
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   9045
      Width           =   1230
   End
   Begin VB.CommandButton cmdMarcar 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Marcar Todos"
      Height          =   330
      Left            =   90
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   8640
      Width           =   1230
   End
   Begin VB.CommandButton cmdDesmarcar 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Desmarcar Todos"
      Height          =   330
      Left            =   1350
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   8640
      Width           =   1455
   End
   Begin VB.CommandButton cmdEtiquetas 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Etiquetas"
      Height          =   870
      Left            =   3825
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   9045
      Width           =   1230
   End
   Begin VB.CommandButton cmdcancel 
      BackColor       =   &H00E0E0E0&
      Caption         =   "ESC-Salir"
      Height          =   870
      Left            =   14085
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   9045
      Width           =   1050
   End
   Begin VB.CommandButton cmdAnadir 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Añadir"
      Height          =   870
      Left            =   60
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   9045
      Width           =   1230
   End
   Begin VB.CommandButton cmdModificar 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Modificar"
      Height          =   870
      Left            =   1305
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   9045
      Width           =   1230
   End
   Begin VB.CommandButton cmdEliminar 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Eliminar"
      Height          =   870
      Left            =   2565
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   9045
      Width           =   1230
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
      Height          =   1455
      Left            =   45
      TabIndex        =   11
      Top             =   540
      Width           =   15090
      Begin VB.TextBox txtb 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   315
         Index           =   1
         Left            =   1305
         TabIndex        =   25
         Top             =   945
         Width           =   4200
      End
      Begin VB.TextBox txtb 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   315
         Index           =   0
         Left            =   1305
         TabIndex        =   0
         Top             =   225
         Width           =   4200
      End
      Begin pryCombo.miCombo cmbResponsable 
         Height          =   330
         Left            =   1305
         TabIndex        =   18
         Top             =   585
         Width           =   5250
         _ExtentX        =   9260
         _ExtentY        =   582
      End
      Begin MSDataListLib.DataCombo cmbTipo 
         Height          =   315
         Left            =   8550
         TabIndex        =   19
         Top             =   225
         Width           =   4530
         _ExtentX        =   7990
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         Style           =   2
         Text            =   ""
      End
      Begin MSDataListLib.DataCombo cmbCentro 
         Height          =   315
         Left            =   8550
         TabIndex        =   20
         Top             =   585
         Width           =   4530
         _ExtentX        =   7990
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         Style           =   2
         Text            =   ""
      End
      Begin VB.CommandButton cmdLimpiar 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Limpiar"
         Height          =   795
         Left            =   13815
         Style           =   1  'Graphical
         TabIndex        =   21
         Top             =   360
         Width           =   915
      End
      Begin MSDataListLib.DataCombo cmbZona 
         Height          =   315
         Left            =   8550
         TabIndex        =   23
         Top             =   945
         Width           =   4530
         _ExtentX        =   7990
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         Style           =   2
         Text            =   ""
      End
      Begin VB.Label Label2 
         BackColor       =   &H00C0C0C0&
         Caption         =   "IP"
         Height          =   195
         Index           =   5
         Left            =   90
         TabIndex        =   26
         Top             =   990
         Width           =   615
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Zona"
         Height          =   195
         Index           =   4
         Left            =   7830
         TabIndex        =   24
         Top             =   990
         Width           =   375
      End
      Begin VB.Label Label2 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Responsable"
         Height          =   195
         Index           =   3
         Left            =   90
         TabIndex        =   17
         Top             =   630
         Width           =   1020
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Centro"
         Height          =   195
         Index           =   2
         Left            =   7830
         TabIndex        =   14
         Top             =   630
         Width           =   465
      End
      Begin VB.Label Label2 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Tipo"
         Height          =   195
         Index           =   1
         Left            =   7830
         TabIndex        =   13
         Top             =   270
         Width           =   615
      End
      Begin VB.Label Label2 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Nombre"
         Height          =   195
         Index           =   0
         Left            =   90
         TabIndex        =   12
         Top             =   270
         Width           =   615
      End
   End
   Begin MSComctlLib.ListView lista 
      Height          =   6585
      Left            =   45
      TabIndex        =   1
      Top             =   2025
      Width           =   15120
      _ExtentX        =   26670
      _ExtentY        =   11615
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
   Begin VB.Label lbltitulo 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "Inventario de Equipos informáticos y Software"
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
      TabIndex        =   16
      Top             =   0
      Width           =   4740
   End
   Begin VB.Label lblsubtitulo 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "Texto"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   90
      TabIndex        =   15
      Top             =   285
      Width           =   405
   End
   Begin VB.Label lblCampos 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      Caption         =   "Pulse sobre la lista para ver el detalle"
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
      Index           =   2
      Left            =   4845
      TabIndex        =   8
      Top             =   8640
      Width           =   3765
   End
   Begin VB.Shape fondo 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   555
      Left            =   0
      Top             =   0
      Width           =   15195
   End
End
Attribute VB_Name = "frmInventario_Listado"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmbZona_Change()
    cargar_lista
End Sub
Private Sub cmdExcel_Click()
   On Error GoTo cmdExcel_Click_Error

    If lista.ListItems.Count = 0 Then Exit Sub
    Dim cadena As String
    Me.MousePointer = vbHourglass
    Dim i As Integer

     Dim XLA As excel.Application
     Dim XLW As excel.Workbook
     Dim XLS As excel.Worksheet
     
     Set XLA = New excel.Application
     Set XLW = XLA.Workbooks.Add
     Set XLS = XLW.Worksheets(1)
     XLW.Worksheets(3).Delete
     XLW.Worksheets(2).Delete
     XLW.Worksheets(1).Name = "Listado de Inventario"

     'Cabecera
     With XLS.Range("A1:I1")
         .HorizontalAlignment = xlCenter
         .VerticalAlignment = xlCenter
         .WrapText = True
         .Orientation = 0
         .AddIndent = False
         .IndentLevel = 0
         .ShrinkToFit = False
         .ReadingOrder = xlContext
         .MergeCells = False
     End With
     With XLS.Range("A1:I1").Interior
         .Pattern = xlSolid
         .PatternColorIndex = xlAutomatic
         .color = &HC0C0FF
     End With
     With XLS.Range("A1:I1").Borders
         .LineStyle = vbSolid
     End With
     
     XLS.Cells(1, 1) = "Codigo"
     XLS.Cells(1, 2) = "Tipo"
     XLS.Cells(1, 3) = "Número"
     XLS.Cells(1, 4) = "Nombre"
     XLS.Cells(1, 5) = "Centro"
     XLS.Cells(1, 6) = "Responsable"
     XLS.Cells(1, 7) = "Zona"
     XLS.Cells(1, 8) = "IP"
     XLS.Cells(1, 9) = "Observaciones"
     XLS.Range("A1:A1").ColumnWidth = 15
     XLS.Range("B1:B1").ColumnWidth = 15
     XLS.Range("C1:C1").ColumnWidth = 15
     XLS.Range("D1:D1").ColumnWidth = 40
     XLS.Range("E1:E1").ColumnWidth = 20
     XLS.Range("F1:F1").ColumnWidth = 30
     XLS.Range("G1:G1").ColumnWidth = 30
     XLS.Range("H1:H1").ColumnWidth = 20
     XLS.Range("I1:I1").ColumnWidth = 100

     For i = 1 To lista.ListItems.Count
         XLS.Cells(i + 1, 1) = ClrStr(lista.ListItems(i).Text, False, True, True)
         XLS.Cells(i + 1, 2) = ClrStr(lista.ListItems(i).SubItems(1), False, True, True)
         XLS.Cells(i + 1, 3) = ClrStr(lista.ListItems(i).SubItems(2), False, True, True)
         XLS.Cells(i + 1, 4) = ClrStr(lista.ListItems(i).SubItems(3), False, True, True)
         XLS.Cells(i + 1, 5) = ClrStr(lista.ListItems(i).SubItems(4), False, True, True)
         XLS.Cells(i + 1, 6) = ClrStr(lista.ListItems(i).SubItems(5), False, True, True)
         XLS.Cells(i + 1, 7) = ClrStr(lista.ListItems(i).SubItems(6), False, True, True)
         XLS.Cells(i + 1, 8) = ClrStr(lista.ListItems(i).SubItems(7), False, True, True)
         XLS.Cells(i + 1, 9) = ClrStr(lista.ListItems(i).SubItems(8), False, True, True)
     Next
     Me.MousePointer = vbNormal
     XLA.Visible = True

   On Error GoTo 0
   Exit Sub

cmdExcel_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure cmdExcel_Click of Formulario frmOferta_Listado"

End Sub
Private Sub chkAgroalimentario_Click()
    cargar_lista
End Sub

Private Sub chkFacturaElectronica_Click()
    cargar_lista
End Sub

Private Sub chkIberia_Click()
    cargar_lista
End Sub

Private Sub chkIntra_Click()
    cargar_lista
End Sub

Private Sub cmbCentro_Change()
    cargar_lista
End Sub

Private Sub cmbResponsable_Change()
    cargar_lista
End Sub

Private Sub chkAirbus_Click()
    cargar_lista
End Sub

Private Sub chkExtranjero_Click()
    cargar_lista
End Sub

Private Sub cmbTipo_change()
    cargar_lista
End Sub

Private Sub cmdduplicar_Click()
    If lista.ListItems.Count = 0 Then Exit Sub
    Dim oInventario As New clsInventario
    Dim ID As Long
    ID = oInventario.duplicar(lista.ListItems(lista.selectedItem.Index))
    If ID = 0 Then
       MsgBox "Error al duplicar los datos.", vbCritical, Err.Description
    Else
       cargar_lista
       frmInventario_Detalle.PK = ID
       frmInventario_Detalle.Show 1
       cargar_lista
    End If
    Set oInventario = Nothing
End Sub

Private Sub cmdAnadir_Click()
    frmInventario_Detalle.PK = 0
    frmInventario_Detalle.Show 1
    cargar_lista
End Sub

Private Sub cmdcancel_Click()
    Unload Me
End Sub

Private Sub cmdDesmarcar_Click()
    Dim i As Integer
    For i = 1 To lista.ListItems.Count
        lista.ListItems(i).Checked = False
    Next
End Sub

Private Sub cmdEliminar_Click()
    If MsgBox("Va a ELIMINAR EL inventario NÚMERO : " & lista.ListItems(lista.selectedItem.Index).Text & ". ¿Esta seguro?", vbExclamation + vbYesNo, "Informacion") = vbYes Then
        Dim oInventario As New clsInventario
        If oInventario.Eliminar(lista.ListItems(lista.selectedItem.Index).Text) = True Then
            cargar_lista
        End If
        Set oInventario = Nothing
    End If

End Sub

Private Sub cmdEtiquetas_Click()
    On Error GoTo fallo
    Dim consulta As String
    Dim registros As String
    Dim generar As Boolean
    generar = False
    Dim i As Integer
    For i = 1 To lista.ListItems.Count
        If lista.ListItems(i).Checked = True Then
            generar = True
            registros = registros & lista.ListItems(i).Text & ","
        End If
    Next
    If generar = False Then
        MsgBox "Marque algún registro para generar las etiquetas.", vbInformation, App.Title
        Exit Sub
    End If
    registros = Left(registros, Len(registros) - 1)
    frmReport.iniciar
    frmReport.informe = "\inventario\rptInventarioEtiqueta"
    frmReport.criterio = "{inventario.ID} IN [" & registros & "] and {tipos.CODIGO} = " & DECODIFICADORA.DECODIFICADORA_INVENTARIO_TIPOS & "  and {zonas.CODIGO} = " & DECODIFICADORA.DECODIFICADORA_INVENTARIO_ZONAS
    frmReport.imprimir = False
    frmReport.pdf = ""
    frmReport.generar
    frmReport.Visible = True
    Exit Sub
fallo:
    MsgBox "Error al generar la etiquetas. " & Err.Description, vbCritical, App.Title
End Sub


Private Sub cmdLimpiar_Click()
    txtb(0) = ""
    cmbTipo.BoundText = ""
    cmbResponsable.Limpiar
    cmbCentro.BoundText = ""
End Sub

'Private Sub cmdImprimir_Click_old()
'    On Error GoTo fallo
'    Dim i As Integer
'    ' Generamos los datos del listado
'    Dim rs As New ADODB.RecordSet
'    rs.Fields.Append "c1", adChar, 5, adFldUpdatable
'    rs.Fields.Append "c2", adChar, 150, adFldUpdatable
'    rs.Fields.Append "c3", adChar, 150, adFldUpdatable
'    rs.Fields.Append "c4", adChar, 150, adFldUpdatable
'    rs.Fields.Append "c5", adChar, 150, adFldUpdatable
'    rs.Open
'    For i = 1 To lista.ListItems.Count
'        rs.AddNew
'        rs("c1") = lista.ListItems(i)
'        If Trim(lista.ListItems(i).SubItems(1)) <> "" Then
'            rs("c2") = lista.ListItems(i).SubItems(1)
'        End If
'        If Trim(lista.ListItems(i).SubItems(2)) <> "" Then
'            rs("c3") = lista.ListItems(i).SubItems(2)
'        End If
'        If Trim(lista.ListItems(i).SubItems(3)) <> "" Then
'            rs("c4") = lista.ListItems(i).SubItems(3)
'        End If
'        If Trim(lista.ListItems(i).SubItems(5)) <> "" Then
'            rs("c5") = lista.ListItems(i).SubItems(5)
'        End If
'        rs.Update
'    Next
'
'    ' Generar Listado
'    Dim Listado As New rptListadoClientes
'    ' Cabecera
'    With Listado.Sections("cabecera")
'        .Controls("titulo").Caption = "Listado de Clientes"
'    End With
'    Set Listado.Sections("cabecera").Controls("logoc").Picture = LoadPicture(ReadINI(App.Path + "\config.ini", "logo", "logo"))
'    'Detalle
'    With Listado.Sections("detalle")
'        .Controls("d1").DataField = rs.Fields("c1").Name
'        .Controls("d2").DataField = rs.Fields("c2").Name
'        .Controls("d3").DataField = rs.Fields("c3").Name
'        .Controls("d4").DataField = rs.Fields("c4").Name
'        .Controls("d5").DataField = rs.Fields("c5").Name
'    End With
'
'    ' Pie de Pagina
'    With Listado.Sections("pie")
'        .Controls("pie1").Caption = "Fecha : " & Format(Date, "dd-mm-yyyy")
'        .Controls("pie2").Caption = "Impreso por : " & USUARIO.getNOMBRE
'    End With
'    Set Listado.DataSource = rs
'    Listado.Caption = "Listado de Clientes"
''    Listado.WindowState = vbMaximized
'    Listado.Show
'    Set rs = Nothing
'    Exit Sub
'fallo:
'    MsgBox "Error al generar el listado de documentos.", vbCritical, Err.Description
'End Sub
'
Private Sub cmdMarcar_Click()
    Dim i As Integer
    For i = 1 To lista.ListItems.Count
        lista.ListItems(i).Checked = True
    Next
End Sub

Private Sub cmdModificar_Click()
    frmInventario_Detalle.PK = lista.ListItems(lista.selectedItem.Index)
    frmInventario_Detalle.Show 1
    actualizar_lista
End Sub

Private Sub cmdResponsable_change()
    cargar_lista
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case 27
            cmdcancel_Click
    End Select
End Sub
Private Sub Form_Load()
    log (Me.Name)
    cargar_botones Me
    Me.Left = 80
    Me.top = 80
    cargar_combos
    With lista.ColumnHeaders
        .Add , , "Codigo", 800, lvwColumnLeft
        .Add , , "Tipo", 1500, lvwColumnCenter
        .Add , , "Número", 1000, lvwColumnCenter
        .Add , , "Nombre", 3300, lvwColumnCenter
        .Add , , "Centro", 1100, lvwColumnCenter
        .Add , , "Responsable", 2100, lvwColumnCenter
        .Add , , "Zona", 2000, lvwColumnCenter
        .Add , , "IP", 1500, lvwColumnCenter
        .Add , , "Observaciones", 1500, lvwColumnCenter
    End With
    cargar_lista
    permisos
End Sub
Private Sub cargar_combos()
    Dim oDeco As New clsDecodificadora
    oDeco.cargar_combo cmbTipo, DECODIFICADORA_INVENTARIO_TIPOS
    oDeco.cargar_combo cmbZona, DECODIFICADORA_INVENTARIO_ZONAS
    Set oDeco = Nothing
    cargar_combo cmbCentro, New clsCentros
    llenar_combo cmbResponsable, New clsEmpleados, 0, frmEmpleados_Gestion, ""
End Sub

Private Sub permisos()
    If USUARIO.getPER_MOD_CLIENTE = False Then
        cmdEliminar.Enabled = False
    End If
End Sub


Private Sub cargar_lista()
    Dim rs As New ADODB.Recordset
    Dim oInventario As New clsInventario
    Set rs = oInventario.Listado(txtb(0), cmbTipo.BoundText, cmbResponsable.getPK_SALIDA, cmbCentro.BoundText, txtb(1), cmbZona.BoundText)
    lista.ListItems.Clear
    If rs.RecordCount <> 0 Then
        Do
            With lista.ListItems.Add(, , Format(rs(0), "0000"))
                .SubItems(1) = rs(1)
                .SubItems(2) = rs(2) ' Número
                .SubItems(3) = rs(3) ' Nombre
                .SubItems(4) = rs(4) ' Centro
                .SubItems(5) = rs(5) ' Responsable
                .SubItems(6) = rs(6) ' Zona
                .SubItems(7) = rs(7) ' ip
                .SubItems(8) = rs(8) ' ob
            End With
            rs.MoveNext
        Loop Until rs.EOF
    End If
    lblsubtitulo = "Total listado : " & lista.ListItems.Count
    Set oInventario = Nothing
    If lista.ListItems.Count > 0 Then
        lista_Click
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
Private Sub lista_Click()
    If lista.ListItems(lista.selectedItem.Index) <> "" Then
      cmdModificar.Enabled = True
      cmdEliminar.Enabled = True
    End If
End Sub
Private Sub lista_DblClick()
    cmdModificar_Click
End Sub
Private Sub actualizar_lista()
    Dim oInventario As New clsInventario
    If oInventario.Carga(CLng(lista.ListItems(lista.selectedItem.Index).Text)) = True Then
        lista.ListItems(lista.selectedItem.Index).SubItems(3) = oInventario.getNOMBRE
        Dim oCentro As New clsCentros
        oCentro.Carga oInventario.getCENTRO_ID
        lista.ListItems(lista.selectedItem.Index).SubItems(4) = oCentro.getNOMBRE
        If oInventario.getUSUARIO_ID = 0 Then
            lista.ListItems(lista.selectedItem.Index).SubItems(5) = ""
        Else
            Dim oUsuario As New clsEmpleados
            oUsuario.cargar oInventario.getUSUARIO_ID
            lista.ListItems(lista.selectedItem.Index).SubItems(5) = oUsuario.getNOMBRE
        End If
        Dim oDeco As New clsDecodificadora
        oDeco.Carga_valor DECODIFICADORA.DECODIFICADORA_INVENTARIO_ZONAS, oInventario.getZONA_ID
        lista.ListItems(lista.selectedItem.Index).SubItems(6) = oDeco.getDESCRIPCION
        lista.ListItems(lista.selectedItem.Index).SubItems(7) = oInventario.getIP
        lista.ListItems(lista.selectedItem.Index).SubItems(8) = oInventario.getOBSERVACIONES
    End If
    Set oInventario = Nothing
    If lista.ListItems.Count > 0 Then
        lista_Click
    End If
End Sub

Private Sub txtb_Change(Index As Integer)
    cargar_lista
End Sub

Private Sub txtb_GotFocus(Index As Integer)
    txtb(Index).BackColor = &H80C0FF
    txtb(Index).SelStart = 0
    txtb(Index).SelLength = Len(txtb(Index))
End Sub
Private Sub txtb_LostFocus(Index As Integer)
    txtb(Index).BackColor = &HFFFFFF
End Sub
