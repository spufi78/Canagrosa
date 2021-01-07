VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form frmBuscarArticulo 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Localizar artículo"
   ClientHeight    =   9060
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   12960
   Icon            =   "frmBuscarArticulo.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9060
   ScaleWidth      =   12960
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdmodificar 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Modificar"
      Height          =   885
      Left            =   1320
      Picture         =   "frmBuscarArticulo.frx":08CA
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   8160
      Width           =   1155
   End
   Begin VB.CommandButton cmdAnadir 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      Caption         =   "&Nuevo"
      Height          =   885
      Left            =   60
      MaskColor       =   &H00E0E0E0&
      Picture         =   "frmBuscarArticulo.frx":1194
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   8160
      Width           =   1215
   End
   Begin VB.CommandButton cmdok 
      BackColor       =   &H00E0E0E0&
      Caption         =   "F10-Aceptar"
      Height          =   855
      Left            =   10380
      Picture         =   "frmBuscarArticulo.frx":1A5E
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   8160
      Width           =   1215
   End
   Begin VB.CommandButton cmdcancel 
      BackColor       =   &H00E0E0E0&
      Caption         =   "ESC-Salir"
      Height          =   855
      Left            =   11640
      Picture         =   "frmBuscarArticulo.frx":2328
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   8160
      Width           =   1245
   End
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   "Criterios de búsqueda de artículos"
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
      Height          =   1470
      Left            =   30
      TabIndex        =   11
      Top             =   45
      Width           =   12855
      Begin VB.CommandButton cmdBuscar 
         BackColor       =   &H00E0E0E0&
         Caption         =   "&Buscar"
         Default         =   -1  'True
         Height          =   915
         Left            =   11430
         Picture         =   "frmBuscarArticulo.frx":2BF2
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   360
         Width           =   1275
      End
      Begin VB.TextBox txtDatos 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   330
         Index           =   1
         Left            =   1350
         TabIndex        =   0
         Top             =   330
         Width           =   3825
      End
      Begin VB.TextBox txtDatos 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   330
         Index           =   0
         Left            =   7095
         TabIndex        =   1
         Top             =   330
         Width           =   4005
      End
      Begin MSDataListLib.DataCombo cmbfamilias 
         Height          =   315
         Left            =   1350
         TabIndex        =   2
         Top             =   705
         Width           =   3825
         _ExtentX        =   6747
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         Style           =   2
         ListField       =   ""
         BoundColumn     =   ""
         Text            =   ""
         Object.DataMember      =   ""
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
      Begin MSDataListLib.DataCombo cmbpro 
         Height          =   315
         Left            =   7095
         TabIndex        =   3
         Top             =   705
         Width           =   4005
         _ExtentX        =   7064
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         Style           =   2
         ListField       =   ""
         BoundColumn     =   ""
         Text            =   ""
         Object.DataMember      =   ""
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
      Begin MSDataListLib.DataCombo cmbSub 
         Height          =   315
         Left            =   1350
         TabIndex        =   4
         Top             =   1080
         Width           =   3825
         _ExtentX        =   6747
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         Style           =   2
         ListField       =   ""
         BoundColumn     =   ""
         Text            =   ""
         Object.DataMember      =   ""
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
      Begin MSDataListLib.DataCombo cmbTipo 
         Height          =   315
         Left            =   7095
         TabIndex        =   5
         Top             =   1080
         Width           =   4005
         _ExtentX        =   7064
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         Style           =   2
         ListField       =   ""
         BoundColumn     =   ""
         Text            =   ""
         Object.DataMember      =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Microsoft Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Subfamilia"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   4
         Left            =   120
         TabIndex        =   18
         Top             =   1110
         Width           =   900
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Tipo"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   3
         Left            =   5940
         TabIndex        =   17
         Top             =   1140
         Width           =   360
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Descripción"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   1
         Left            =   120
         TabIndex        =   16
         Top             =   360
         Width           =   990
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Proveedor"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   6
         Left            =   5940
         TabIndex        =   14
         Top             =   765
         Width           =   885
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Familia"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   0
         Left            =   120
         TabIndex        =   13
         Top             =   735
         Width           =   600
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Código"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   2
         Left            =   5940
         TabIndex        =   12
         Top             =   390
         Width           =   600
      End
   End
   Begin MSComctlLib.ListView lista 
      Height          =   6570
      Left            =   45
      TabIndex        =   15
      Top             =   1545
      Width           =   12855
      _ExtentX        =   22675
      _ExtentY        =   11589
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      AllowReorder    =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   14609914
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
Attribute VB_Name = "frmBuscarArticulo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdAnadir_Click()
    garticulo = 0
    frmAnadirArticulo.Show 1
End Sub

Private Sub cmdBuscar_Click()
    On Error GoTo fallo
    Dim consulta As String
    Dim proveedores As String
    Dim familias As String
    Dim subfamilias As String
    Dim codigo As String
    Dim tipo As String
    proveedores = ""
    familias = ""
    subfamilias = ""
    codigo = ""
    tipo = ""
    If cmbfamilias.Text <> "" Then
        familias = " AND art.FAMILIA_ID = " & cmbfamilias.BoundText
    End If
    If cmbSub.Text <> "" Then
        subfamilias = " AND art.subFAMILIA_ID = " & cmbSub.BoundText
    End If
    If cmbpro.Text <> "" Then
        proveedores = " AND art.PROVEEDOR_ID = " & cmbpro.BoundText
    End If
    If cmbTipo.Text <> "" Then
        tipo = " AND art.TIPO_ARTICULO_ID = " & cmbTipo.BoundText
    End If
    If txtDatos(0) <> "" Then
'        If Len(txtDatos(0)) > 4 Or IsNumeric(txtDatos(0)) = False Then
'            codigo = "AND art.EAN like '%" & txtDatos(0) & "%'"
'        Else
        If IsNumeric(txtDatos(0)) Then

            codigo = "AND art.id_ARTICULO = " & txtDatos(0)
        End If
    End If
    Dim rs As New ADODB.Recordset
    consulta = "SELECT art.id_articulo, " & _
               "       art.descripcion, " & _
               "       fam.nombre, " & _
               "       sub.nombre, " & _
               "       art.precio_venta, " & _
               "       art.stock, " & _
               "       art.ean " & _
               "  FROM articulos as art, " & _
               "       familias as fam, " & _
               "       subfamilias as sub " & _
               " WHERE art.familia_id = fam.id_familia " & _
               "   and art.subfamilia_id = sub.id_subfamilia " & _
               "   and art.descripcion like '%" & txtDatos(1) & "%'" & _
               codigo & _
               familias & _
               subfamilias & _
               proveedores & _
               tipo & _
               " ORDER BY art.descripcion"
    lista.ListItems.Clear
    Me.MousePointer = 11
    Set rs = datos_bd(consulta)
    If rs.RecordCount >= 1 Then
        While Not rs.EOF
            With lista.ListItems.Add(, , Format(rs.Fields(0), "0000000"))
'                .SubItems(1) = rs.Fields(6)
                .SubItems(1) = Format(rs.Fields(0), "0000000")
                .SubItems(2) = rs.Fields(1)
                .SubItems(3) = rs.Fields(2)
                .SubItems(4) = rs.Fields(3)
                .SubItems(5) = Format(rs.Fields(4), "currency")
                .SubItems(6) = rs.Fields(5)
            End With
            rs.MoveNext
        Wend
        lista.SetFocus
    Else
        MsgBox "No existen artículos con ese criterio.", vbInformation, App.Title
    End If
    Set rs = Nothing
    Me.MousePointer = 0
    Exit Sub
fallo:
    Me.MousePointer = 0
    MsgBox "Se ha producido un error al buscar los artículos.", vbCritical, Err.Description
End Sub

Private Sub cmdcancel_Click()
    garticulo = 0
    Unload Me
End Sub

Private Sub cmdModificar_Click()
    If lista.ListItems.Count > 0 Then
        garticulo = lista.ListItems(lista.SelectedItem.Index)
        frmAnadirArticulo.Show 1
        actualizar_lista
        garticulo = 0
        lista.SetFocus
    End If
End Sub

Private Sub cmdok_Click()
    If lista.ListItems.Count > 0 Then
        garticulo = lista.ListItems(lista.SelectedItem.Index)
        Unload Me
    Else
        garticulo = 0
    End If
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
     Case 27
        cmdcancel_Click
     Case 121 ' F10
        cmdok_Click
    End Select
End Sub

Private Sub Form_Load()
    log (Me.Name)
    Call cabecera
    Call cargar_combos
End Sub

Private Sub lista_DblClick()
    cmdok_Click
'    If lista.ListItems.Count > 0 Then
'        garticulo = lista.ListItems(lista.SelectedItem.Index)
'        frmAnadirArticulo.Show 1
'        actualizar_lista
'        garticulo = 0
'        lista.SetFocus
'    End If
End Sub
Private Sub txtDatos_GotFocus(Index As Integer)
    txtDatos(Index).BackColor = &H80C0FF
    txtDatos(Index).SelStart = 0
    txtDatos(Index).SelLength = Len(txtDatos(Index))
End Sub
Private Sub txtdatos_LostFocus(Index As Integer)
    txtDatos(Index).BackColor = vbWhite
End Sub

Public Sub cabecera()
    With lista.ColumnHeaders.Add(, , "Codigo", 1, lvwColumnLeft)
        .Tag = "Codigo"
    End With
    With lista.ColumnHeaders.Add(, , "Codigo", 1200, lvwColumnCenter)
        .Tag = "Codigo"
    End With
    With lista.ColumnHeaders.Add(, , "Descripcion", 4000, lvwColumnLeft)
        .Tag = "Descripcion"
    End With
    With lista.ColumnHeaders.Add(, , "Familia", 2600, lvwColumnLeft)
        .Tag = "Familia"
    End With
    With lista.ColumnHeaders.Add(, , "SubFamilia", 2600, lvwColumnLeft)
        .Tag = "SubFamilia"
    End With
    With lista.ColumnHeaders.Add(, , "Precio", 1100, lvwColumnRight)
        .Tag = "Precio"
    End With
    With lista.ColumnHeaders.Add(, , "Existencias", 1000, lvwColumnCenter)
        .Tag = "Existencias"
    End With
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
   On Error GoTo cargar_combos_Error

    cargar_combo cmbfamilias, New clsFamilias
    cargar_combo cmbpro, New clsProveedor
    cargar_combo cmbTipo, New clsArticulos_Tipos
    cmbfamilias.BoundText = 1
   On Error GoTo 0
   Exit Sub

cargar_combos_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure cargar_combos of Formulario frmBuscarArticulo"
    
End Sub
Private Sub cmbfamilias_Change()
    If cmbfamilias.Text <> "" Then
     If IsNumeric(cmbfamilias.BoundText) Then
        Dim osubfamilia As New clsSubfamilias
        Set cmbSub.RowSource = osubfamilia.Listado_por_familia(CInt(cmbfamilias.BoundText))  'recorset devuelto por la funcion
        cmbSub.ListField = "nombre" 'campo que veo
        cmbSub.DataField = "nombre" 'campo asociado
        cmbSub.BoundColumn = "id_subfamilia" 'lo que realmente envia
        Set osubfamilia = Nothing
     End If
    End If
End Sub
Public Sub actualizar_lista()
    Dim oart As New clsArticulo
    Dim ofamilia As New clsFamilias
    Dim osubfamilia As New clsSubfamilias
    If oart.Cargar(lista.ListItems(lista.SelectedItem.Index)) = True Then
        lista.ListItems(lista.SelectedItem.Index).SubItems(1) = oart.getEAN
        lista.ListItems(lista.SelectedItem.Index).SubItems(2) = oart.getDESCRIPCION
        If oart.getFAMILIA_ID <> 0 Then
            ofamilia.Carga (oart.getFAMILIA_ID)
            lista.ListItems(lista.SelectedItem.Index).SubItems(3) = ofamilia.getNOMBRE
        End If
        If oart.getSUBFAMILIA_ID <> 0 Then
            osubfamilia.Carga (oart.getSUBFAMILIA_ID)
            lista.ListItems(lista.SelectedItem.Index).SubItems(4) = osubfamilia.getNOMBRE
        End If
        lista.ListItems(lista.SelectedItem.Index).SubItems(5) = Format(oart.getPRECIO_VENTA, "currency")
        lista.ListItems(lista.SelectedItem.Index).SubItems(6) = oart.getSTOCK
    End If
    Set oart = Nothing
End Sub

