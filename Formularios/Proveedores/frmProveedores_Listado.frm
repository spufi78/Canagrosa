VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form frmProveedores_Listado 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Listado de Proveedores"
   ClientHeight    =   9645
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   14085
   Icon            =   "frmProveedores_Listado.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   9645
   ScaleWidth      =   14085
   Begin VB.CommandButton cmdEtiquetas 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Etiquetas"
      Height          =   870
      Left            =   10845
      Style           =   1  'Graphical
      TabIndex        =   30
      Top             =   8730
      Width           =   1050
   End
   Begin VB.CommandButton cmdLi15 
      BackColor       =   &H00E0E0E0&
      Caption         =   "LI - 15 (Subcontrata)"
      Height          =   870
      Left            =   9765
      Picture         =   "frmProveedores_Listado.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   29
      Top             =   8730
      Width           =   1050
   End
   Begin VB.CommandButton cmdCorreo 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Correo"
      Height          =   870
      Left            =   4320
      Picture         =   "frmProveedores_Listado.frx":0BD4
      Style           =   1  'Graphical
      TabIndex        =   28
      Top             =   8730
      Width           =   1050
   End
   Begin VB.CommandButton cmdAdjuntos 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Facturas"
      Height          =   870
      Left            =   5400
      Style           =   1  'Graphical
      TabIndex        =   27
      ToolTipText     =   "Enviar informe por E-mail"
      Top             =   8730
      Width           =   1065
   End
   Begin VB.CommandButton cmdEvaluacion 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Evaluación"
      Height          =   870
      Left            =   6495
      Picture         =   "frmProveedores_Listado.frx":149E
      Style           =   1  'Graphical
      TabIndex        =   26
      Top             =   8730
      Width           =   1065
   End
   Begin VB.CommandButton cmdAnular 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Anular"
      Height          =   870
      Left            =   3240
      Picture         =   "frmProveedores_Listado.frx":1D68
      Style           =   1  'Graphical
      TabIndex        =   24
      Top             =   8730
      Width           =   1050
   End
   Begin VB.CommandButton cmdLI14 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Lista LI - 14"
      Height          =   870
      Left            =   8685
      Picture         =   "frmProveedores_Listado.frx":2632
      Style           =   1  'Graphical
      TabIndex        =   23
      Top             =   8730
      Width           =   1050
   End
   Begin VB.CommandButton cmdFicha 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Ficha"
      Height          =   870
      Left            =   7590
      Picture         =   "frmProveedores_Listado.frx":2EFC
      Style           =   1  'Graphical
      TabIndex        =   19
      Top             =   8730
      Width           =   1050
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
      TabIndex        =   6
      Top             =   585
      Width           =   13965
      Begin VB.CheckBox chkExtra 
         BackColor       =   &H00C0C0C0&
         Caption         =   "ExtraComunitarios"
         Height          =   240
         Left            =   12150
         TabIndex        =   33
         Top             =   855
         Width           =   1695
      End
      Begin VB.CheckBox chkIntra 
         BackColor       =   &H00C0C0C0&
         Caption         =   "IntraComunitarios"
         Height          =   240
         Left            =   12150
         TabIndex        =   32
         Top             =   585
         Width           =   1695
      End
      Begin VB.CheckBox chkFormadores 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Sólo Formadores"
         Height          =   420
         Left            =   10215
         TabIndex        =   31
         Top             =   630
         Width           =   1695
      End
      Begin VB.CheckBox chkAnulados 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Mostrar proveedores anulados"
         Height          =   420
         Left            =   8235
         TabIndex        =   25
         Top             =   660
         Width           =   1785
      End
      Begin VB.CheckBox chkNA 
         BackColor       =   &H00C0C0C0&
         Caption         =   "No Asignados a Equipos ni Reactivos"
         Height          =   375
         Left            =   4365
         TabIndex        =   22
         Top             =   675
         Width           =   1860
      End
      Begin VB.CheckBox chkAR 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Asignados a Reactivos"
         Height          =   375
         Left            =   3015
         TabIndex        =   21
         Top             =   675
         Width           =   1320
      End
      Begin VB.CheckBox chkAE 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Asignados a Equipos"
         Height          =   375
         Left            =   1485
         TabIndex        =   20
         Top             =   675
         Width           =   1230
      End
      Begin VB.CheckBox chkSS 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Mostrar proveedores sin servicio asignado"
         Height          =   420
         Left            =   6300
         TabIndex        =   18
         Top             =   645
         Width           =   1770
      End
      Begin VB.CheckBox chkMostrarSoloSubcontratas 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Mostrar sólo subcontratas"
         Height          =   375
         Left            =   90
         TabIndex        =   17
         Top             =   675
         Width           =   1275
      End
      Begin MSDataListLib.DataCombo cmbServicio 
         Height          =   315
         Left            =   8505
         TabIndex        =   16
         Top             =   225
         Width           =   5190
         _ExtentX        =   9155
         _ExtentY        =   556
         _Version        =   393216
         Style           =   2
         Text            =   ""
      End
      Begin VB.TextBox txtb 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   315
         Index           =   0
         Left            =   870
         TabIndex        =   9
         Top             =   225
         Width           =   1860
      End
      Begin VB.TextBox txtb 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   315
         Index           =   1
         Left            =   3330
         TabIndex        =   8
         Top             =   225
         Width           =   1455
      End
      Begin VB.TextBox txtb 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   315
         Index           =   2
         Left            =   5895
         TabIndex        =   7
         Top             =   225
         Width           =   1635
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Servicio"
         Height          =   195
         Index           =   3
         Left            =   7785
         TabIndex        =   15
         Top             =   270
         Width           =   570
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
      Begin VB.Label Label2 
         BackColor       =   &H00C0C0C0&
         Caption         =   "CIF"
         Height          =   195
         Index           =   1
         Left            =   2970
         TabIndex        =   11
         Top             =   270
         Width           =   615
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Telefono"
         Height          =   195
         Index           =   2
         Left            =   5085
         TabIndex        =   10
         Top             =   270
         Width           =   630
      End
   End
   Begin VB.CommandButton cmdImprimir 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Listado"
      Height          =   870
      Left            =   11925
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   8730
      Width           =   1050
   End
   Begin VB.CommandButton cmdEliminar 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Eliminar"
      Height          =   870
      Left            =   2175
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   8730
      Width           =   1050
   End
   Begin VB.CommandButton cmdModificar 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Modificar"
      Height          =   870
      Left            =   1130
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   8730
      Width           =   1050
   End
   Begin VB.CommandButton cmdAnadir 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Añadir"
      Height          =   870
      Left            =   60
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   8730
      Width           =   1050
   End
   Begin VB.CommandButton cmdcancel 
      BackColor       =   &H00E0E0E0&
      Caption         =   "ESC-Salir"
      Height          =   870
      Left            =   13005
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   8730
      Width           =   1050
   End
   Begin MSComctlLib.ListView lista 
      Height          =   6945
      Left            =   45
      TabIndex        =   0
      Top             =   1755
      Width           =   13995
      _ExtentX        =   24686
      _ExtentY        =   12250
      View            =   3
      LabelEdit       =   1
      MultiSelect     =   -1  'True
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
      Caption         =   "Listado de Proveedores"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   90
      TabIndex        =   14
      Top             =   315
      Width           =   1680
   End
   Begin VB.Label lbltitulo 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "Listado de Proveedores"
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
      TabIndex        =   13
      Top             =   45
      Width           =   2520
   End
   Begin VB.Shape fondo 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   600
      Left            =   0
      Top             =   0
      Width           =   14160
   End
End
Attribute VB_Name = "frmProveedores_Listado"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Sub chkExtra_Click()
    cargar_lista
End Sub

Private Sub chkFormadores_Click()
    cargar_lista
End Sub

Private Sub chkIntra_Click()
    cargar_lista
End Sub

Private Sub cmdAdjuntos_Click()
    If lista.ListItems.Count > 0 Then
'M1257-I
        frmProveedores_Facturas.TOBJETO = 0
        frmProveedores_Facturas.COBJETO = 0
'M1257-F
        frmProveedores_Facturas.PK = lista.ListItems(lista.selectedItem.Index)
        frmProveedores_Facturas.Show 1
    End If
End Sub


Private Sub cmdCorreo_Click()
    Dim i As Integer
    Dim ASUNTO As String
    Dim CORREO As String
    Dim ref As String
    Dim adjunto As String
'    adjunto = "\\servidor\canagrosa\Cuestionario COD POC 0405 ANEXO VII.docx"
    Dim rs As ADODB.Recordset
    Set rs = datos_bd("select * from correo WHERE ID = 11 ")
    If rs.RecordCount > 0 Then
        ASUNTO = rs("asunto")
        CORREO = rs("CORREO")
        If Not IsNull(rs("ADJUNTO")) Then
            adjunto = rs("ADJUNTO")
        End If
    End If
    For i = 1 To lista.ListItems.Count
        If lista.ListItems(i).Selected = True Then
            If Trim(lista.ListItems(i).SubItems(4)) <> "" Then
                genera_correo lista.ListItems(i).SubItems(4), ASUNTO, CORREO, adjunto, Me.hdc, True
            End If
        End If
    Next

End Sub

Private Sub cmdEtiquetas_Click()
    On Error GoTo fallo
    Dim consulta As String
    Dim proveedores As String
    Dim generar As Boolean
    generar = False
    Dim i As Integer
    If lista.ListItems.Count = 0 Then Exit Sub
    proveedores = lista.ListItems(lista.selectedItem.Index).Text
'    For i = 1 To lista.ListItems.Count
'        If lista.ListItems(i).Checked = True Then
'            generar = True
'            proveedores = proveedores & lista.ListItems(i).Text & ","
'        End If
'    Next
'    If generar = False Then
'        MsgBox "Marque algún proveedor para generar las etiquetas.", vbInformation, App.Title
'        Exit Sub
'    End If
'    proveedores = Left(proveedores, Len(proveedores) - 1)
    frmReport.iniciar
    frmReport.informe = "\proveedores\rptEtiquetaSobre"
    frmReport.criterio = "{proveedores.ID_PROVEEDOR} IN [" & proveedores & "]"
    frmReport.imprimir = False
    frmReport.pdf = ""
    frmReport.generar
    frmReport.visible = True
    Exit Sub
fallo:
    MsgBox "Error al generar la etiquetas. " & Err.Description, vbCritical, App.Title

End Sub

Private Sub cmdEvaluacion_Click()
    If lista.ListItems.Count = 0 Then Exit Sub
    frmProveedores_Evaluacion.PK = lista.ListItems(lista.selectedItem.Index)
    frmProveedores_Evaluacion.Show 1
End Sub

Private Sub chkAE_Click()
    If chkAE.Value = Checked Then
        chkAR.Value = Unchecked
        chkNA.Value = Unchecked
        chkSS.Value = Unchecked
    End If
    cargar_lista
End Sub

Private Sub chkAnulados_Click()
    cargar_lista

End Sub

Private Sub chkAR_Click()
    If chkAR.Value = Checked Then
        chkAE.Value = Unchecked
        chkNA.Value = Unchecked
        chkSS.Value = Unchecked
    End If
    cargar_lista
End Sub

Private Sub chkNA_Click()
    If chkNA.Value = Checked Then
        chkAR.Value = Unchecked
        chkAE.Value = Unchecked
        chkSS.Value = Unchecked
    End If
    cargar_lista
End Sub

Private Sub chkSS_Click()
    If chkSS.Value = Checked Then
        chkAR.Value = Unchecked
        chkNA.Value = Unchecked
        chkAE.Value = Unchecked
    End If
    cargar_lista
End Sub

Private Sub cmbServicio_Change()
    cargar_lista
End Sub

Private Sub cmdAnular_Click()
    If MsgBox("Va a ANULAR al proveedor " & lista.ListItems(lista.selectedItem.Index).SubItems(1) & ". ¿Esta seguro?", vbExclamation + vbYesNo, "Informacion") = vbYes Then
        Dim oProveedor As New clsProveedor
        If oProveedor.Anular(lista.ListItems(lista.selectedItem.Index)) = True Then
            cargar_lista
        End If
        Set oProveedor = Nothing
    End If
End Sub

Private Sub cmdFicha_Click()
    Dim oprov As New clsProveedor
    oprov.ImprimirFicha lista.ListItems(lista.selectedItem.Index)
    Set oprov = Nothing
End Sub

Private Sub cmdImprimir_Click()
    Dim oprov As New clsProveedor
    oprov.ImprimirListadoCompleto txtb(0), txtb(1), txtb(2), chkMostrarSoloSubcontratas.Value, chkAnulados.Value
    Set oprov = Nothing
End Sub

'E0200-I
Private Sub chkMostrarSoloSubcontratas_Click()
    cargar_lista
End Sub
'E0200-F

Private Sub cmdAnadir_Click()
    'E0061-I
    ' Se cambia gproveedor por PK
    'gproveedor = 0
    frmProveedores_Detalle.PK = 0
    'E0061-F
    
    frmProveedores_Detalle.Show 1
    cargar_lista
End Sub

Private Sub cmdcancel_Click()
    Unload Me
End Sub

Private Sub cmdEliminar_Click()
    If MsgBox("Va a ELIMINAR al proveedor " & lista.ListItems(lista.selectedItem.Index).SubItems(1) & ". ¿Esta seguro?", vbExclamation + vbYesNo, "Informacion") = vbYes Then
        Dim oProveedor As New clsProveedor
        oProveedor.setID_PROVEEDOR = lista.ListItems(lista.selectedItem.Index)
        If oProveedor.Eliminar = True Then
            cargar_lista
        End If
        Set oProveedor = Nothing
    End If
End Sub

Private Sub cmdLI14_Click()
    Dim oprov As New clsProveedor
    oprov.ImprimirListado
    Set oprov = Nothing
End Sub

Private Sub cmdLi15_Click()
    Dim oprov As New clsProveedor
    oprov.ImprimirListadoLI15
    Set oprov = Nothing
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
'        rs.Update
'    Next
'
'    ' Generar Listado
'    Dim Listado As New rptListado
'    ' Cabecera
'    With Listado.Sections("cabecera")
'        .Controls("titulo").Caption = "Listado de Proveedores"
'    End With
'    Set Listado.Sections("cabecera").Controls("logoc").Picture = LoadPicture(ReadINI(App.Path + "\config.ini", "logo", "logo"))
'    'Detalle
'    With Listado.Sections("detalle")
'        .Controls("d1").DataField = rs.Fields("c1").Name
'        .Controls("d2").DataField = rs.Fields("c2").Name
'        .Controls("d3").DataField = rs.Fields("c3").Name
'        .Controls("d4").DataField = rs.Fields("c4").Name
'    End With
'
'    ' Pie de Pagina
'    With Listado.Sections("pie")
'        .Controls("pie1").Caption = "Fecha : " & Format(Date, "dd-mm-yyyy")
'        .Controls("pie2").Caption = "Impreso por : " & Usuario.getNOMBRE
'    End With
'    Set Listado.DataSource = rs
'    Listado.Caption = "Listado de Proveedores"
''    Listado.WindowState = vbMaximized
'    Listado.Show
'    Set rs = Nothing
'    Exit Sub
'fallo:
'    MsgBox "Error al generar el listado de documentos.", vbCritical, Err.Description
'End Sub

Private Sub cmdModificar_Click()
    'E0063-I
    ' se cambia gproveedor por PK
    'gproveedor = lista.ListItems(lista.SelectedItem.Index)
    frmProveedores_Detalle.PK = lista.ListItems(lista.selectedItem.Index)
    'E0063-F
    
    frmProveedores_Detalle.Show 1
    actualizar_lista
    'cargar_lista

    'E0064-I
    ' se cambia gproveedor por PK
    'gproveedor = 0
    ' frmProveedores_Detalle.PK = 0
    'E0064-F
End Sub

Private Sub Command1_Click()
    Dim i As Integer
    Dim c As String
    Dim ref As String
    Dim rs As ADODB.Recordset
    Set rs = datos_bd("select correo from correo")
    If rs.RecordCount > 0 Then
           c = rs(0)
    End If
    ref = "REQUISITOS MEDIOAMBIENTALES"
    For i = 1 To lista.ListItems.Count
        If lista.ListItems(i).Selected = True Then
            genera_correo lista.ListItems(i).SubItems(4), ref, c, "", Me.hdc, True
        End If
    Next

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
    Me.top = 100
    Me.Left = 100
    Dim oDeco As New clsDecodificadora
    oDeco.cargar_combo cmbServicio, DECODIFICADORA.PROVEEDORES_SERVICIOS
    With lista.ColumnHeaders
        .Add , , "Codigo", 600, lvwColumnLeft
        .Add , , "Nombre", 4300, lvwColumnLeft
        .Add , , "Direccion", 4300, lvwColumnLeft
        .Add , , "Telefono", 1300, lvwColumnCenter
        .Add , , "Email", 2700, lvwColumnLeft
    End With
    permisos
    cargar_lista
End Sub
Private Sub permisos()
    If USUARIO.getPER_TESORERIA_FP = False Then
        cmdAdjuntos.visible = False
    End If
End Sub
Public Sub cargar_lista()
    Dim rs As New ADODB.Recordset
    Dim ocli As New clsProveedor
   On Error GoTo cargar_lista_Error

    Me.MousePointer = 11
    Set rs = ocli.Listado(chkMostrarSoloSubcontratas.Value, txtb(0), txtb(1), txtb(2), cmbServicio.BoundText, chkSS.Value, chkAE.Value, chkAR.Value, chkNA.Value, chkAnulados.Value, chkFormadores.Value, chkIntra.Value, chkExtra.Value)
    lista.ListItems.Clear
    lblsubtitulo = "Registros encontrados : " & rs.RecordCount
    If rs.RecordCount <> 0 Then
        Do
           With lista.ListItems.Add(, , Format(rs("id_proveedor"), "0000"))
            .SubItems(1) = rs("nombre")
            If IsNull(rs("direccion")) = False Then
                .SubItems(2) = rs("direccion")
            End If
            If IsNull(rs("telefono")) = False Then
                .SubItems(3) = rs("telefono")
            End If
            .SubItems(4) = rs("email")
           End With
           rs.MoveNext
        Loop Until rs.EOF
    End If
    Set ocli = Nothing
    If lista.ListItems.Count > 0 Then
        lista_Click
    End If
    Me.MousePointer = 0

   On Error GoTo 0
   Exit Sub

cargar_lista_Error:
    Me.MousePointer = 0

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure cargar_lista of Formulario frmProveedores_Listado"
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
    If lista.ListItems.Count = 0 Then Exit Sub
    If lista.ListItems(lista.selectedItem.Index) <> "" Then
      cmdModificar.Enabled = True
      cmdEliminar.Enabled = True
    End If
End Sub

Private Sub lista_DblClick()
    cmdModificar_Click
End Sub

Public Sub actualizar_lista()
    Dim oPro As New clsProveedor
    'E0065-I
    ' se cambia gproveedor por PK
    'If ocli.Carga(CLng(gproveedor)) = True Then
    If oPro.Carga(CLng(lista.ListItems(lista.selectedItem.Index).Text)) = True Then
    'E0065-F
'        lista.ListItems(lista.SelectedItem.Index).Text = gproveedor
        lista.ListItems(lista.selectedItem.Index).SubItems(1) = oPro.getNOMBRE
        lista.ListItems(lista.selectedItem.Index).SubItems(2) = oPro.getDIRECCION
        lista.ListItems(lista.selectedItem.Index).SubItems(3) = oPro.getTELEFONO
    End If
    Set oPro = Nothing
    If lista.ListItems.Count > 0 Then
        lista_Click
    End If
End Sub

Private Sub txtb_Change(Index As Integer)
    cargar_lista
End Sub
