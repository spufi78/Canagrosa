VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmREX_Inventario_Listado 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Listado de Inventarios de Reactivos Externos / Productos Controlados"
   ClientHeight    =   8415
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11670
   Icon            =   "frmREX_Inventario_Listado.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   8415
   ScaleWidth      =   11670
   Begin VB.CommandButton cmdImprimir 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Imprimir"
      Height          =   870
      Left            =   9450
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   7470
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
      Height          =   780
      Left            =   45
      TabIndex        =   7
      Top             =   630
      Width           =   11580
      Begin VB.TextBox txtfiltro 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   330
         Index           =   0
         Left            =   990
         TabIndex        =   8
         Top             =   270
         Width           =   1230
      End
      Begin MSComCtl2.DTPicker fdesde 
         Height          =   330
         Left            =   3015
         TabIndex        =   11
         Top             =   270
         Width           =   1320
         _ExtentX        =   2328
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
         Format          =   51970049
         CurrentDate     =   38002
      End
      Begin MSComCtl2.DTPicker fhasta 
         Height          =   330
         Left            =   4905
         TabIndex        =   12
         Top             =   270
         Width           =   1275
         _ExtentX        =   2249
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
         Format          =   51970049
         CurrentDate     =   38002
      End
      Begin MSDataListLib.DataCombo cmbUsuario 
         Height          =   315
         Left            =   6930
         TabIndex        =   13
         Top             =   270
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         Style           =   2
         Text            =   ""
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
      Begin MSDataListLib.DataCombo cmbCentro 
         Bindings        =   "frmREX_Inventario_Listado.frx":1272
         Height          =   315
         Left            =   9630
         TabIndex        =   17
         Top             =   270
         Width           =   1875
         _ExtentX        =   3307
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
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Centro"
         Height          =   195
         Index           =   22
         Left            =   9090
         TabIndex        =   18
         Top             =   315
         Width           =   465
      End
      Begin VB.Label lblCampos 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "hasta"
         Height          =   195
         Index           =   2
         Left            =   4455
         TabIndex        =   15
         Top             =   315
         Width           =   405
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Usuario"
         Height          =   195
         Index           =   8
         Left            =   6345
         TabIndex        =   14
         Top             =   315
         Width           =   540
      End
      Begin VB.Label Label2 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Descripción"
         Height          =   285
         Index           =   3
         Left            =   90
         TabIndex        =   10
         Top             =   315
         Width           =   870
      End
      Begin VB.Label Label2 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Fecha"
         Height          =   195
         Index           =   1
         Left            =   2385
         TabIndex        =   9
         Top             =   315
         Width           =   600
      End
   End
   Begin VB.CommandButton cmdcancel 
      BackColor       =   &H00E0E0E0&
      Caption         =   "ESC-Salir"
      Height          =   870
      Left            =   10575
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   7485
      Width           =   1050
   End
   Begin VB.CommandButton cmdAnadir 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Añadir"
      Height          =   870
      Left            =   60
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   7470
      Width           =   1050
   End
   Begin VB.CommandButton cmdModificar 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Modificar"
      Height          =   870
      Left            =   1125
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   7470
      Width           =   1050
   End
   Begin VB.CommandButton cmdEliminar 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Eliminar"
      Height          =   870
      Left            =   2205
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   7470
      Width           =   1050
   End
   Begin MSComctlLib.ListView lista 
      Height          =   5970
      Left            =   45
      TabIndex        =   0
      Top             =   1425
      Width           =   11595
      _ExtentX        =   20452
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
      Caption         =   "Listado de Inventarios de Reactivos Externos / Productos Controlados"
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
      TabIndex        =   6
      Top             =   45
      Width           =   7275
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
      TabIndex        =   5
      Top             =   315
      Width           =   2775
   End
   Begin VB.Shape fondo 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   600
      Left            =   0
      Top             =   0
      Width           =   12240
   End
End
Attribute VB_Name = "frmREX_Inventario_Listado"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Function listado_botes_por_inventario(ID_INVENTARIO As Long)
Dim consulta As String

 consulta = "SELECT be.id_bote_ex, " & _
               "       tb.codigo, " & _
               "       tr.nombre, " & _
               "       be.fecha_recepcion, " & _
               "       be.fecha_apertura, " & _
               "       be.fecha_fin, " & _
               "       be.fecha_caducidad, " & _
               "       be.tipo_bote_ex_id, " & _
               "       be.LOTE, " & _
               "       tb.precio, " & _
               "       tb.cantidad, tb.tipo_m_referencia_id, be.numero,be.codigo " & _
               " FROM BOTES_EX be, " & _
               "      TIPOS_BOTE_EX tb, " & _
               "      TIPOS_REACTIVO_EX tr " & _
               " WHERE be.tipo_bote_ex_id = tb.id_tipo_bote_ex " & _
               " AND tb.tipo_reactivo_ex_id = tr.id_tipo_reactivo_ex " & _
               " AND be.id_bote_ex in (select bote_ex_id from rex_inventarios_botes where inventario_id = " & ID_INVENTARIO & ")" & _
               " ORDER BY be.id_bote_ex asc"
               
        Set listado_botes_por_inventario = datos_bd(consulta)
            
End Function

Private Sub cmbCentro_Change()
    cargar_lista
End Sub

Private Sub cmbUsuario_Change()
    cargar_lista
End Sub

Private Sub cmdAnadir_Click()
    frmREX_Inventario.PK = 0
    frmREX_Inventario.Show 1
    cargar_lista
End Sub

Private Sub cmdcancel_Click()
    Unload Me
End Sub

Private Sub cmdEliminar_Click()
    If MsgBox("Va a ELIMINAR el inventario " & lista.ListItems(lista.selectedItem.Index).SubItems(1) & ". ¿Esta seguro?", vbExclamation + vbYesNo, "Informacion") = vbYes Then
        Dim oInventario As New clsRex_inventarios
        If oInventario.Eliminar(lista.ListItems(lista.selectedItem.Index)) = True Then
            lista.ListItems.Remove lista.selectedItem.Index
            If lista.ListItems.Count > 0 Then
                If lista.selectedItem.Index < lista.ListItems.Count Then
                    Set lista.selectedItem = lista.ListItems(lista.selectedItem.Index)
                End If
            End If
        End If
        Set oInventario = Nothing
    End If
End Sub

'Private Sub cmdImprimir_Click_old()
'On Error GoTo fallo
'
'    If lista.SelectedItem.Index <= 0 Then Exit Sub
'
'    Dim i As Integer
'    Dim total As Currency
'    ' Generamos los datos del listado
'    Dim rs As New ADODB.RecordSet
'    Dim rs_botes As New ADODB.RecordSet
'
'    Set rs_botes = listado_botes_por_inventario(lista.ListItems(lista.SelectedItem.Index))
'
'    rs.Fields.Append "c1", adChar, 15, adFldUpdatable
'    rs.Fields.Append "c2", adChar, 150, adFldUpdatable
'    rs.Fields.Append "c3", adChar, 150, adFldUpdatable
'    rs.Fields.Append "c4", adChar, 150, adFldUpdatable
'    rs.Fields.Append "c5", adChar, 15, adFldUpdatable ' Precio
'    rs.Open
'    i = 1
'    If rs_botes.RecordCount <> 0 Then
'        rs_botes.MoveFirst
'        While Not rs_botes.EOF
'            rs.AddNew
'            'If Trim(rs_botes(13)) = "" Then
'                rs("c1") = rs_botes(0)
'            'Else
'            '    rs("c1") = CStr(rs_botes(13)) & "-" & Format(rs_botes(12), "000") & "-" & Format(rs_botes(3), "yy")
'            'End If
'            If Trim(CStr(rs_botes(2))) <> "" Then
'                rs("c2") = CStr(rs_botes(2))
'            End If
'            If Trim(CStr(rs_botes(8))) <> "" Then
'                rs("c3") = CStr(rs_botes(8))
'            End If
'            If Trim(CStr(rs_botes(10))) <> "" Then
'                rs("c4") = CStr(rs_botes(10))
'            End If
'            If Trim(CStr(rs_botes(9))) <> "" Then
'                rs("c5") = CStr(rs_botes(9))
'                total = total + rs_botes(9)
'            End If
'            rs.Update
'            rs_botes.MoveNext
'        Wend
'    End If
'
'
'    ' Generar Listado
'    Dim Listado As New rptListadoReactivos
'    ' Cabecera
'    With Listado.Sections("cabecera")
'        .Controls("titulo").Caption = "Inventario de Reactivos: " & lista.ListItems(lista.SelectedItem.Index).SubItems(1)
'        .Controls("etiqueta4").Caption = "Número"
'        .Controls("etiqueta5").Caption = "Reactivo"
'
'        .Controls("etiqueta10").Caption = "Lote"
'        .Controls("etiqueta11").Caption = "Cantidad"
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
'    ' Pie de Pagina
'    With Listado.Sections("pie")
'        .Controls("pie1").Caption = "Fecha : " & Format(Date, "dd-mm-yyyy")
'        .Controls("pie2").Caption = "Impreso por : " & USUARIO.getNOMBRE
'    End With
'    With Listado.Sections("totales")
'        .Controls("LBLT1").Caption = Format(total, "CURRENCY")
'    End With
'    Set Listado.DataSource = rs
'    Listado.Caption = "Listado de Botes de Reactivos"
'    Listado.Show
'    Set rs = Nothing
'    Exit Sub
'fallo:
'    MsgBox "Error al generar el listado: " & Err.Number & ": " & Err.Description, vbCritical, Err.Description
'End Sub
'
Private Sub cmdImprimir_Click()
Dim ocli As New clsRex_inventarios

    If lista.ListItems.Count = 0 Then Exit Sub
    If lista.selectedItem.Index <= 0 Then Exit Sub

    ocli.Imprimir_Listado lista.ListItems(lista.selectedItem.Index).SubItems(1), lista.ListItems(lista.selectedItem.Index)

Set ocli = Nothing
End Sub

Private Sub cmdModificar_Click()
    If lista.ListItems.Count > 0 Then
        frmREX_Inventario.PK = lista.ListItems(lista.selectedItem.Index)
        frmREX_Inventario.Show 1
        actualizar_lista
    End If
End Sub

Private Sub fdesde_Change()
cargar_lista
End Sub

Private Sub fhasta_Change()
cargar_lista
End Sub

Private Sub Form_Activate()
    Me.SetFocus
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
    Me.Left = 100
    Me.top = 100
    cabecera
    cargar_combos
    fdesde = Date - 1800
    fhasta = Date
    cargar_lista
End Sub
Private Sub cargar_lista()
    Dim rs As ADODB.Recordset
    Dim oInventario As New clsRex_inventarios
    Set rs = oInventario.Listado(txtfiltro(0), fdesde.Value, fhasta.Value, cmbUsuario.BoundText, cmbCentro.BoundText)
    lista.ListItems.Clear
    lbltitulo(1) = "En la lista existen un total de " & rs.RecordCount & " registros."
    If rs.RecordCount <> 0 Then
        Do
           With lista.ListItems.Add(, , Format(rs(0), "0000"))
            .SubItems(1) = rs(1)
            .SubItems(2) = Format(rs(2), "dd-mm-yyyy")
            .SubItems(3) = rs(3)
            .SubItems(4) = rs(4)
            .SubItems(5) = rs(5)
           End With
           rs.MoveNext
        Loop Until rs.EOF
    End If
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
    If lista.ListItems.Count > 0 Then
      cmdModificar.Enabled = True
      cmdEliminar.Enabled = True
    Else
      cmdModificar.Enabled = False
      cmdEliminar.Enabled = False
    End If
End Sub

Private Sub lista_DblClick()
    If lista.ListItems.Count > 0 Then
        cmdModificar_Click
    End If
End Sub

Private Sub actualizar_lista()
    Dim rs As ADODB.Recordset
    Dim oInventario As New clsRex_inventarios
    Set rs = oInventario.Listado_PK(lista.ListItems(lista.selectedItem.Index).Text)
    With lista.ListItems(lista.selectedItem.Index)
        .SubItems(1) = rs(1)
        .SubItems(2) = Format(rs(2), "dd-mm-yyyy")
        .SubItems(3) = rs(3)
        .SubItems(4) = rs(4)
        .SubItems(5) = rs(5)
    End With
    Set oInventario = Nothing
End Sub

Private Sub txtfiltro_Change(Index As Integer)
    cargar_lista
End Sub

Private Sub cabecera()
    With lista.ColumnHeaders
        .Add , , "Inventario", 600, lvwColumnLeft
        .Add , , "Descripción", 5500, lvwColumnLeft
        .Add , , "Fecha", 1300, lvwColumnCenter
        .Add , , "Centro", 1200, lvwColumnCenter
        .Add , , "Usuario", 1200, lvwColumnCenter
        .Add , , "Estado", 1200, lvwColumnCenter
    End With
End Sub

Private Sub cargar_combos()
    cargar_combo cmbUsuario, New clsUsuarios
    cargar_combo cmbCentro, New clsCentros
End Sub
