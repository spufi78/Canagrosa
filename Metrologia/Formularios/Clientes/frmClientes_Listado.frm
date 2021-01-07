VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmClientes_Listado 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Listado de Clientes"
   ClientHeight    =   8925
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   13110
   Icon            =   "frmClientes_Listado.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   8925
   ScaleWidth      =   13110
   Begin VB.CommandButton cmdObras 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Obras"
      Height          =   885
      Left            =   7200
      Picture         =   "frmClientes_Listado.frx":1CFA
      Style           =   1  'Graphical
      TabIndex        =   19
      Top             =   7995
      Width           =   1215
   End
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   "Filtro de búsqueda"
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
      Height          =   1095
      Left            =   60
      TabIndex        =   14
      Top             =   375
      Width           =   13005
      Begin VB.CheckBox chkanulados 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Mostrar anulados"
         Height          =   345
         Left            =   135
         TabIndex        =   4
         Top             =   675
         Width           =   2775
      End
      Begin VB.TextBox txtDatos 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   0
         Left            =   4320
         TabIndex        =   1
         Top             =   300
         Width           =   1965
      End
      Begin VB.TextBox txtDatos 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   1
         Left            =   1230
         TabIndex        =   0
         Top             =   300
         Width           =   2145
      End
      Begin VB.TextBox txtDatos 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   2
         Left            =   7560
         TabIndex        =   2
         Top             =   315
         Width           =   1905
      End
      Begin VB.TextBox txtDatos 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   3
         Left            =   10890
         TabIndex        =   3
         Top             =   315
         Width           =   1965
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "N.I.F."
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   2
         Left            =   3645
         TabIndex        =   18
         Top             =   360
         Width           =   555
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Teléfono"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   0
         Left            =   6480
         TabIndex        =   17
         Top             =   360
         Width           =   855
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Población"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   6
         Left            =   9840
         TabIndex        =   16
         Top             =   360
         Width           =   930
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Nombre"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   1
         Left            =   135
         TabIndex        =   15
         Top             =   360
         Width           =   735
      End
   End
   Begin VB.CommandButton cmdAnular 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Anular"
      Height          =   885
      Left            =   3615
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   7995
      Width           =   1155
   End
   Begin VB.CommandButton cmdDocumentos 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Documentos"
      Height          =   885
      Left            =   8460
      Picture         =   "frmClientes_Listado.frx":26E4
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   7995
      Width           =   1215
   End
   Begin VB.CommandButton cmdImprimir 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Listado"
      Height          =   885
      Left            =   6005
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   7995
      Width           =   1155
   End
   Begin VB.CommandButton cmdBuscar 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Buscar"
      Height          =   885
      Left            =   4810
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   7995
      Width           =   1155
   End
   Begin VB.CommandButton cmdeliminar 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Eliminar"
      Height          =   885
      Left            =   2420
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   7995
      Width           =   1155
   End
   Begin VB.CommandButton cmdmodificar 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Modificar"
      Height          =   885
      Left            =   1225
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   7995
      Width           =   1155
   End
   Begin VB.CommandButton cmdcancel 
      BackColor       =   &H00E0E0E0&
      Caption         =   "ESC-Salir"
      Height          =   885
      Left            =   11910
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   7995
      Width           =   1155
   End
   Begin VB.CommandButton cmdanadir 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Añadir"
      Height          =   885
      Left            =   30
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   7995
      Width           =   1155
   End
   Begin MSComctlLib.ListView lista 
      Height          =   6420
      Left            =   60
      TabIndex        =   13
      Top             =   1515
      Width           =   13005
      _ExtentX        =   22939
      _ExtentY        =   11324
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
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Caption         =   "Listado de Clientes"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   345
      Index           =   3
      Left            =   0
      TabIndex        =   20
      Top             =   0
      Width           =   13035
   End
End
Attribute VB_Name = "frmClientes_Listado"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub chkanulados_Click()
    cargar_lista
End Sub

Private Sub cmdAnadir_Click()
    frmClientes.pk = 0
    frmClientes.Show 1
    cargar_lista
    lista.SetFocus
End Sub

Private Sub cmdAnular_Click()
    If lista.ListItems.Count > 0 Then
        Dim pos As Integer
        Dim cliente As Integer
        Dim Mensaje As String
        If lista.ListItems(lista.SelectedItem.Index).SubItems(5) = 0 Then
            Mensaje = "Va a ANULAR al Cliente "
        Else
            Mensaje = "Va a RESTABLECER al Cliente "
        End If
        If MsgBox(Mensaje & lista.ListItems(lista.SelectedItem.Index).SubItems(1) & ". ¿Esta seguro?", vbExclamation + vbYesNo, "Informacion") = vbYes Then
            Dim ocliente As New clsCliente
            ocliente.setID_CLIENTE = lista.ListItems(lista.SelectedItem.Index)
            If lista.ListItems(lista.SelectedItem.Index).SubItems(5) = 0 Then
                ocliente.anular_cliente
            Else
                ocliente.restablecer_cliente
            End If
            cargar_lista
            Set ocliente = Nothing
        End If
        lista.SetFocus
    End If
End Sub

Private Sub cmdBuscar_Click()
    gcliente = 0
    frmBuscarCliente.Show 1
    If gcliente <> 0 Then
        Dim i As Integer
        For i = 1 To lista.ListItems.Count
            If lista.ListItems(i) = gcliente Then
                lista.ListItems(i).Selected = True
                lista.ListItems(i).EnsureVisible
                lista.SetFocus
                Exit Sub
            End If
        Next
    End If
End Sub
Private Sub cmdcancel_Click()
    Unload Me
End Sub

Private Sub cmdDocumentos_Click()
    If lista.ListItems.Count > 0 Then
        gcliente = lista.ListItems(lista.SelectedItem.Index)
        Dim ofrm As New frmListadoDocumentos
        ofrm.Show
        Set ofrm = Nothing
    End If
End Sub

Private Sub cmdEliminar_Click()
    If lista.ListItems.Count > 0 Then
        Dim pos As Integer
        Dim cliente As Integer
        If MsgBox("Va a ELIMINAR al Cliente " & lista.ListItems(lista.SelectedItem.Index).SubItems(1) & ". ¿Esta seguro?", vbExclamation + vbYesNo, "Informacion") = vbYes Then
            Dim ocliente As New clsCliente
            ocliente.setID_CLIENTE = lista.ListItems(lista.SelectedItem.Index)
            If ocliente.eliminar_cliente = True Then
                cargar_lista
            End If
            Set ocliente = Nothing
        End If
        lista.SetFocus
    End If
End Sub

Private Sub cmdImprimir_Click()
    Dim FILTRO As String
    If txtDatos(1) <> "" Then
        FILTRO = FILTRO & " {clientes.NOMBRE} like '*" & txtDatos(1) & "*'"
    End If
    If txtDatos(0) <> "" Then
        If FILTRO <> "" Then
            FILTRO = FILTRO & " AND "
        End If
        FILTRO = FILTRO & " {clientes.CIF} like '*" & txtDatos(0) & "*'"
    End If
    If txtDatos(2) <> "" Then
        If FILTRO <> "" Then
            FILTRO = FILTRO & " AND "
        End If
        FILTRO = FILTRO & " {clientes.TELEFONO} like '*" & txtDatos(2) & "*'"
    End If
    If txtDatos(3) <> "" Then
        If FILTRO <> "" Then
            FILTRO = FILTRO & " AND "
        End If
        FILTRO = FILTRO & " {municipios.NOMBRE} like '*" & txtDatos(3) & "*'"
    End If
    With frmReport
        .iniciar
        .CRITERIO = FILTRO
        .informe = "rptClientes"
        .imprimir = False
        .generar
        .Show 1
    End With
    Unload frmReport
End Sub

Private Sub cmdModificar_Click()
    If USUARIO.getPER_3 = 0 Then
        Exit Sub
    End If
    If lista.ListItems.Count > 0 Then
        frmClientes.pk = lista.ListItems(lista.SelectedItem.Index)
        frmClientes.Show 1
        actualizar_lista
        lista.SetFocus
    End If
End Sub


Private Sub cmdObras_Click()
    If lista.ListItems.Count > 0 Then
        Dim ofrm As New frmObras_Listado
        ofrm.PK_CLIENTE = lista.ListItems(lista.SelectedItem.Index)
        ofrm.Show
        Set ofrm = Nothing
    End If
End Sub

Private Sub Form_Activate()
    Me.SetFocus
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case 27 ' esc
            Unload Me
    End Select
End Sub

Private Sub Form_Load()
    log (Me.Name)
    cargar_botones Me
    Me.Left = 100
    Me.Top = 100
    With lista.ColumnHeaders.Add(, , "Codigo", 1000, lvwColumnLeft)
        .Tag = "Codigo"
    End With
    With lista.ColumnHeaders.Add(, , "Nombre", 3500, lvwColumnLeft)
        .Tag = "Nombre"
    End With
    With lista.ColumnHeaders.Add(, , "Direccion", 3900, lvwColumnLeft)
        .Tag = "Direccion"
    End With
    With lista.ColumnHeaders.Add(, , "Provincia", 2200, lvwColumnCenter)
        .Tag = "Provincia"
    End With
    With lista.ColumnHeaders.Add(, , "Poblacion", 2100, lvwColumnCenter)
        .Tag = "Telefono"
    End With
    With lista.ColumnHeaders.Add(, , "ANULADO", 1, lvwColumnCenter)
        .Tag = "ANULADO"
    End With
    cargar_lista
    permisos
End Sub

Public Sub cargar_lista()
    Dim rs As New ADODB.Recordset
    Dim ocli As New clsCliente
    Set rs = ocli.Listado(txtDatos(1), txtDatos(0), txtDatos(2), txtDatos(3), chkanulados.Value)
    lista.ListItems.Clear
    If rs.RecordCount <> 0 Then
        Do
           With lista.ListItems.Add(, , Format(rs(0), "000"))
            .SubItems(1) = rs(1)
            .SubItems(2) = rs(2)
            If Not IsNull(rs(3)) Then
                .SubItems(3) = rs(3)
            End If
            If Not IsNull(rs(4)) Then
                .SubItems(4) = rs(4)
            End If
            If Not IsNull(rs(5)) Then
                .SubItems(5) = rs(5)
            End If
           End With
           rs.MoveNext
        Loop Until rs.EOF
    End If
    Set ocli = Nothing
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
        If lista.ListItems(lista.SelectedItem.Index) <> "" Then
          cmdmodificar.Enabled = True
          cmdeliminar.Enabled = True
        End If
        If lista.ListItems(lista.SelectedItem.Index).SubItems(5) = 0 Then
            cmdAnular.Caption = "Anular"
        Else
            cmdAnular.Caption = "Restablecer"
        End If
        permisos
    End If
End Sub

Private Sub lista_DblClick()
    cmdModificar_Click
End Sub

Public Sub actualizar_lista()
    Dim ocli As New clsCliente
    If ocli.CargaCliente(lista.ListItems(lista.SelectedItem.Index).Text) = True Then
        lista.ListItems(lista.SelectedItem.Index).SubItems(1) = ocli.getNOMBRE
        lista.ListItems(lista.SelectedItem.Index).SubItems(2) = ocli.getDIRECCION
        Dim oProvincia As New clsProvincias
        oProvincia.Carga (ocli.getPROVINCIA_ID)
        lista.ListItems(lista.SelectedItem.Index).SubItems(3) = oProvincia.getNOMBRE
        Dim oMunicipio As New clsMunicipios
        oMunicipio.Cargar ocli.getMUNICIPIO_ID
        lista.ListItems(lista.SelectedItem.Index).SubItems(4) = oMunicipio.getNOMBRE
        lista.ListItems(lista.SelectedItem.Index).SubItems(5) = ocli.getANULADO
    End If
    Set ocli = Nothing
    If lista.ListItems.Count > 0 Then
        lista_Click
    End If
End Sub
Private Sub lista_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
       cmdModificar_Click
    End If
End Sub
Public Sub permisos()
    If USUARIO.getPER_1 = 0 Then
        cmdImprimir.Enabled = False
    End If
    If USUARIO.getPER_2 = 0 Then
        cmdanadir.Enabled = False
    End If
    If USUARIO.getPER_3 = 0 Then
        cmdmodificar.Enabled = False
    End If
    If USUARIO.getPER_4 = 0 Then
        cmdeliminar.Enabled = False
    End If
End Sub

Private Sub txtDatos_Change(Index As Integer)
    cargar_lista
End Sub

Private Sub txtDatos_GotFocus(Index As Integer)
    txtDatos(Index).BackColor = &H80C0FF
    txtDatos(Index).SelStart = 0
    txtDatos(Index).SelLength = Len(txtDatos(Index))
End Sub

Private Sub txtDatos_LostFocus(Index As Integer)
    txtDatos(Index).BackColor = vbWhite
End Sub
