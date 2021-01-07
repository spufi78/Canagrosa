VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{CEA3B7F6-2847-4E5E-A551-DB7A62489D44}#46.0#0"; "miCombo.ocx"
Begin VB.Form frmREX_Botes_Listado 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Listado de Botes de Reactivos Externos / Productos Controlados"
   ClientHeight    =   9120
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   12240
   Icon            =   "frmREX_Botes_Listado.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   9120
   ScaleWidth      =   12240
   Begin VB.CommandButton cmdOk 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Recuperar"
      Height          =   870
      Left            =   3285
      Style           =   1  'Graphical
      TabIndex        =   21
      Top             =   8235
      Width           =   1050
   End
   Begin VB.CommandButton cmbTipos 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Tipos"
      Height          =   870
      Left            =   5445
      Picture         =   "frmREX_Botes_Listado.frx":1272
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   8235
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
      Height          =   1410
      Left            =   45
      TabIndex        =   11
      Top             =   765
      Width           =   12120
      Begin VB.CheckBox chkanulados 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Anulados"
         Height          =   240
         Left            =   10215
         TabIndex        =   18
         Top             =   810
         Width           =   1320
      End
      Begin VB.TextBox txtfiltro 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   330
         Index           =   1
         Left            =   4770
         TabIndex        =   1
         Top             =   225
         Width           =   1995
      End
      Begin VB.TextBox txtfiltro 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   330
         Index           =   0
         Left            =   1620
         TabIndex        =   0
         Top             =   225
         Width           =   1860
      End
      Begin MSDataListLib.DataCombo cmbTipo 
         Height          =   315
         Left            =   7785
         TabIndex        =   2
         Top             =   225
         Width           =   3735
         _ExtentX        =   6588
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
      Begin pryCombo.miCombo cmbProveedores 
         Height          =   330
         Left            =   1620
         TabIndex        =   16
         Top             =   630
         Width           =   8070
         _ExtentX        =   14235
         _ExtentY        =   582
      End
      Begin pryCombo.miCombo cmbResponsable 
         Height          =   330
         Left            =   1620
         TabIndex        =   19
         Top             =   990
         Width           =   8070
         _ExtentX        =   14235
         _ExtentY        =   582
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Responsable"
         Height          =   195
         Index           =   1
         Left            =   135
         TabIndex        =   20
         Top             =   1035
         Width           =   930
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Proveedor"
         Height          =   195
         Index           =   0
         Left            =   135
         TabIndex        =   17
         Top             =   675
         Width           =   735
      End
      Begin VB.Label Label2 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Tipo"
         Height          =   195
         Index           =   2
         Left            =   7290
         TabIndex        =   14
         Top             =   315
         Width           =   510
      End
      Begin VB.Label Label2 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Código"
         Height          =   195
         Index           =   1
         Left            =   4095
         TabIndex        =   13
         Top             =   315
         Width           =   600
      End
      Begin VB.Label Label2 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Sustancia/Material"
         Height          =   285
         Index           =   0
         Left            =   135
         TabIndex        =   12
         Top             =   315
         Width           =   1635
      End
   End
   Begin VB.CommandButton cmdImprimir 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Listado"
      Height          =   870
      Left            =   4365
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   8235
      Width           =   1050
   End
   Begin VB.CommandButton cmdEliminar 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Anular"
      Height          =   870
      Left            =   2210
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   8235
      Width           =   1050
   End
   Begin VB.CommandButton cmdModificar 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Modificar"
      Height          =   870
      Left            =   1135
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   8235
      Width           =   1050
   End
   Begin VB.CommandButton cmdAnadir 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Añadir"
      Height          =   870
      Left            =   60
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   8235
      Width           =   1050
   End
   Begin VB.CommandButton cmdcancel 
      BackColor       =   &H00E0E0E0&
      Caption         =   "ESC-Salir"
      Height          =   870
      Left            =   11115
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   8220
      Width           =   1050
   End
   Begin MSComctlLib.ListView lista 
      Height          =   5970
      Left            =   45
      TabIndex        =   8
      Top             =   2205
      Width           =   12120
      _ExtentX        =   21378
      _ExtentY        =   10530
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
   Begin VB.Label lbltitulo 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "En la lista existen un total de 0 registros"
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   1
      Left            =   90
      TabIndex        =   10
      Top             =   360
      Width           =   2775
   End
   Begin VB.Image imagen 
      Height          =   480
      Left            =   11610
      Picture         =   "frmREX_Botes_Listado.frx":1B3C
      Top             =   45
      Width           =   480
   End
   Begin VB.Label lbltitulo 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "Listado de Botes de Reactivos Externos / Productos Controlados"
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
      Index           =   0
      Left            =   90
      TabIndex        =   9
      Top             =   45
      Width           =   6735
   End
   Begin VB.Shape fondo 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   735
      Left            =   0
      Top             =   0
      Width           =   12240
   End
End
Attribute VB_Name = "frmREX_Botes_Listado"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public PK_TIPO_REACTIVO_ID As Long

Private Sub cmbResponsable_Change()
    cargar_lista
End Sub

Private Sub cmdImprimir_Click()
Dim otbe As New clsTipos_bote_ex

    otbe.Imprimir_Listado txtFiltro(0).Text, txtFiltro(1).Text, getDataComboSel(cmbTipo), cmbProveedores.getPK_SALIDA, (chkAnulados.Value = vbChecked)

Set otbe = Nothing
End Sub

Private Sub chkAnulados_Click()
    cargar_lista
    habilitar_controles
End Sub

Private Sub cmbProveedores_change()
    cargar_lista
End Sub

Private Sub cmbTipo_change()
    cargar_lista
End Sub

Private Sub cmbTipos_Click()
    Dim oform As New frmDecodificadora
    oform.CODIGO = DECODIFICADORA.REX_TIPOS
    oform.Show
End Sub

Private Sub cmdAnadir_Click()
'    gbotereactivoex = 0
    frmREX_Bote.PK = 0
    frmREX_Bote.Show 1
    cargar_lista
End Sub
Private Sub cmdcancel_Click()
    Unload Me
End Sub

Private Sub cmdEliminar_Click()
    If MsgBox("Va a ELIMINAR el bote " & lista.ListItems(lista.selectedItem.Index).SubItems(1) & ". ¿Esta seguro?", vbExclamation + vbYesNo, "Informacion") = vbYes Then
        Dim obe As New clsTipos_bote_ex
        If obe.Eliminar(lista.ListItems(lista.selectedItem.Index)) = True Then
            cargar_lista
        End If
        Set obe = Nothing
    End If
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
'    rs.Fields.Append "c5", adChar, 15, adFldUpdatable
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
'        If Trim(lista.ListItems(i).SubItems(4)) <> "" Then
'            rs("c5") = lista.ListItems(i).SubItems(4)
'        End If
'        rs.Update
'    Next
'    ' Generar Listado
'    Dim Listado As New rptListado5
'    ' Cabecera
'    With Listado.Sections("cabecera")
'        .Controls("titulo").Caption = "Listado de Reactivos Externos/Productos Controlados"
'        .Controls("etiqueta4").Caption = "ID"
'        .Controls("etiqueta5").Caption = "Código"
'        .Controls("etiqueta10").Caption = "Reactivo"
'        .Controls("etiqueta11").Caption = "Proveedor"
'        .Controls("etiqueta1").Caption = "Precio"
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
'    Listado.Sections("detalle").Controls("d5").Alignment = 1
'
'    ' Pie de Pagina
'    With Listado.Sections("pie")
'        .Controls("pie1").Caption = "Fecha : " & Format(Date, "dd-mm-yyyy")
'        .Controls("pie2").Caption = "Impreso por : " & USUARIO.getNOMBRE
'    End With
'    Set Listado.DataSource = rs
'    Listado.Caption = "Listado de Botes Externos"
''    Listado.WindowState = vbMaximized
'    Listado.Show
'    Set rs = Nothing
'    Exit Sub
'fallo:
'    MsgBox "Error al generar el listado.", vbCritical, Err.Description
'End Sub
'
Private Sub cmdModificar_Click()
'    gbotereactivoex = lista.ListItems(lista.SelectedItem.Index)
    frmREX_Bote.PK = lista.ListItems(lista.selectedItem.Index)
    frmREX_Bote.Show 1
    actualizar_lista
'    gbotereactivoex = 0
End Sub

'M1144-I
Private Sub cmdok_Click()
    If MsgBox("Va a RECUPERAR el bote " & lista.ListItems(lista.selectedItem.Index).SubItems(1) & ". ¿Esta seguro?", vbExclamation + vbYesNo, "Informacion") = vbYes Then
        Dim obe As New clsTipos_bote_ex
        If obe.Recuperar(lista.ListItems(lista.selectedItem.Index)) = True Then
            cargar_lista
        End If
        Set obe = Nothing
    End If
End Sub
'M1144-F

Private Sub Form_Activate()
    Me.SetFocus
    If PK_TIPO_REACTIVO_ID > 0 Then
        cargar_lista
    End If
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
    cargar_combos
    If PK_TIPO_REACTIVO_ID = 0 Then
        Me.Left = 100
        Me.top = 100
    End If
    cabecera
    cargar_lista
End Sub

Private Sub cargar_lista()
    lista.ListItems.Clear
    Dim rs As New ADODB.Recordset
    Dim otbe As New clsTipos_bote_ex
    If cmbResponsable.getTEXTO = "" Then
        Set rs = otbe.Listado(txtFiltro(0), txtFiltro(1), cmbTipo.BoundText, cmbProveedores.getPK_SALIDA, PK_TIPO_REACTIVO_ID, chkAnulados.Value, 0)
    Else
        Set rs = otbe.Listado(txtFiltro(0), txtFiltro(1), cmbTipo.BoundText, cmbProveedores.getPK_SALIDA, PK_TIPO_REACTIVO_ID, chkAnulados.Value, cmbResponsable.getPK_SALIDA)
    End If
    PK_TIPO_REACTIVO_ID = 0
    lbltitulo(1) = "En la lista existen un total de " & rs.RecordCount & " registros."
    If rs.RecordCount <> 0 Then
        Do
           With lista.ListItems.Add(, , Format(rs(0), "0000"))
            .SubItems(1) = rs(1)
            .SubItems(2) = rs(2)
            .SubItems(3) = rs(3)
            .SubItems(4) = rs(4)
            .SubItems(5) = moneda(rs(5))
           End With
           rs.MoveNext
        Loop Until rs.EOF
    End If
    Set otbe = Nothing
    lista_Click
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
    If lista.ListItems.Count > 0 Then
      cmdModificar.Enabled = True
      habilitar_controles
    Else
      cmdModificar.Enabled = False
      habilitar_controles
    End If
End Sub

Private Sub lista_DblClick()
    If lista.ListItems.Count > 0 Then
        cmdModificar_Click
    End If
End Sub

Private Sub actualizar_lista()
    Dim rs As New ADODB.Recordset
    Dim otbe As New clsTipos_bote_ex
    Set rs = otbe.Listado_ID(lista.ListItems(lista.selectedItem.Index).Text)
    If rs.RecordCount <> 0 Then
        Do
           With lista.ListItems(lista.selectedItem.Index)
            .SubItems(1) = rs(1)
            .SubItems(2) = rs(2)
            .SubItems(3) = rs(3)
            .SubItems(4) = rs(4)
            .SubItems(5) = moneda(rs(5))
           End With
           rs.MoveNext
        Loop Until rs.EOF
    End If
    lista_Click
End Sub

Private Sub txtfiltro_Change(Index As Integer)
    cargar_lista
End Sub
Private Sub cargar_combos()
'    Cargar_Combo cmbtipo, New clsTipos_m_referencia
    Dim oDeco As New clsDecodificadora
    oDeco.cargar_combo cmbTipo, DECODIFICADORA.REX_TIPOS
    llenar_combo cmbResponsable, New clsUsuarios, 0, frmUsuarios, ""
    ' Cargar Proveedores que tienes tipos de botes
    Dim consulta As String
    Dim conn As ADODB.Connection
    If CrearConexionGlobal(conn, "", "") = True Then ' CONECTAR EL RS
        consulta = "SELECT DISTINCT B.ID_PROVEEDOR , B.NOMBRE " & _
                   "  FROM TIPOS_BOTE_EX A, PROVEEDORES B " & _
                   " WHERE A.PROVEEDOR_ID = B.ID_PROVEEDOR "
        With cmbProveedores
            .setCONN = conn
            .setFK_CAMPO = ""
            .setFK_VALOR = 0
            .setTABLA = "PROVEEDORES"
            .setDESCRIPCION = "Proveedores"
            .setPK = "B.ID_PROVEEDOR"
            .setCAMPO = "NOMBRE"
            .setMUESTRA_DETALLE = True
            .setQUERY = consulta
            .setFILTRO = ""
            Set .FORMULARIO = frmProveedores_Detalle
        End With
    End If
End Sub
Private Sub cabecera()
    With lista.ColumnHeaders.Add(, , "Id", 600, lvwColumnLeft)
        .Tag = "Id"
    End With
    With lista.ColumnHeaders.Add(, , "Codigo", 1300, lvwColumnLeft)
        .Tag = "Codigo"
    End With
    With lista.ColumnHeaders.Add(, , "Sustancia/Material", 3600, lvwColumnLeft)
        .Tag = "Tipo de Reactivo"
    End With
    With lista.ColumnHeaders.Add(, , "Proveedor", 3600, lvwColumnLeft)
        .Tag = "Proveedor"
    End With
    With lista.ColumnHeaders.Add(, , "Tipo", 1500, lvwColumnCenter)
        .Tag = "Tipo"
    End With
    With lista.ColumnHeaders.Add(, , "Precio", 1000, lvwColumnRight)
        .Tag = "Precio"
    End With
End Sub
Private Sub habilitar_controles()
    If chkAnulados.Value = 0 Then
       cmdEliminar.Enabled = True
       cmdok.Enabled = False
    Else
       cmdEliminar.Enabled = False
       cmdok.Enabled = True
    End If
End Sub
