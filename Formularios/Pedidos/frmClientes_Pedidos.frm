VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmClientes_Pedidos 
   BackColor       =   &H00C0C0C0&
   Caption         =   "Pedidos de Cliente"
   ClientHeight    =   9735
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   13935
   Icon            =   "frmClientes_Pedidos.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   9735
   ScaleWidth      =   13935
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame2 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Filtrar por"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   690
      Left            =   0
      TabIndex        =   26
      Top             =   540
      Width           =   13905
      Begin VB.TextBox txtFiltro 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   315
         Index           =   1
         Left            =   10530
         MaxLength       =   100
         TabIndex        =   2
         Top             =   225
         Width           =   2445
      End
      Begin VB.TextBox txtFiltro 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   315
         Index           =   0
         Left            =   5895
         MaxLength       =   30
         TabIndex        =   1
         Top             =   225
         Width           =   1905
      End
      Begin MSDataListLib.DataCombo cmbFiltroTipo 
         Height          =   315
         Left            =   675
         TabIndex        =   0
         Top             =   225
         Width           =   3870
         _ExtentX        =   6826
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         Style           =   2
         Text            =   ""
      End
      Begin VB.Label Label3 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Tipo"
         Height          =   240
         Left            =   180
         TabIndex        =   29
         Top             =   270
         Width           =   420
      End
      Begin VB.Label lblCampos 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Descripción"
         Height          =   240
         Index           =   6
         Left            =   9315
         TabIndex        =   28
         Top             =   270
         Width           =   1185
      End
      Begin VB.Label lblCampos 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Código"
         Height          =   240
         Index           =   5
         Left            =   4995
         TabIndex        =   27
         Top             =   270
         Width           =   945
      End
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Archivos Adjuntos (Adjuntar despues de crear el pedido)"
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
      Height          =   2100
      Left            =   6930
      TabIndex        =   24
      Top             =   6615
      Width           =   6945
      Begin VB.CommandButton cmdAdjuntar 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Adjuntar y eliminar"
         Height          =   960
         Left            =   5040
         Picture         =   "frmClientes_Pedidos.frx":3AFA
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   675
         Width           =   1620
      End
      Begin MSComctlLib.ListView listaAdjuntos 
         Height          =   1815
         Left            =   135
         TabIndex        =   9
         Top             =   225
         Width           =   4815
         _ExtentX        =   8493
         _ExtentY        =   3201
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         AllowReorder    =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   16777215
         BorderStyle     =   1
         Appearance      =   0
         NumItems        =   0
      End
   End
   Begin VB.CheckBox chkvigor 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Mostrar solo pedidos en vigor"
      Height          =   255
      Left            =   5850
      TabIndex        =   15
      Top             =   9060
      Value           =   1  'Checked
      Width           =   2655
   End
   Begin VB.CommandButton cmdDetalle 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Detalle"
      Height          =   870
      Left            =   3270
      Picture         =   "frmClientes_Pedidos.frx":43C4
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   8805
      Width           =   1050
   End
   Begin VB.CommandButton cmdEliminar 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Eliminar"
      Height          =   870
      Left            =   2190
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   8805
      Width           =   1050
   End
   Begin VB.CommandButton cmdModificar 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Modificar"
      Height          =   870
      Left            =   1110
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   8805
      Width           =   1050
   End
   Begin VB.CommandButton cmdAnadir 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Añadir"
      Height          =   870
      Left            =   45
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   8805
      Width           =   1050
   End
   Begin VB.CommandButton cmdSalir 
      BackColor       =   &H00E0E0E0&
      Caption         =   "ESC-Salir"
      Height          =   870
      Left            =   12825
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   8775
      Width           =   1050
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Datos del pedido"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2100
      Left            =   45
      TabIndex        =   18
      Top             =   6615
      Width           =   6795
      Begin VB.CheckBox chkRevisadoConforme 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Pedido Revisado y Conforme"
         Height          =   240
         Left            =   3600
         TabIndex        =   31
         Top             =   1710
         Width           =   2490
      End
      Begin VB.TextBox txtDatos 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H000000FF&
         Height          =   315
         Index           =   0
         Left            =   1230
         MaxLength       =   30
         TabIndex        =   8
         Top             =   1695
         Width           =   1575
      End
      Begin VB.TextBox txtDatos 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   315
         Index           =   2
         Left            =   1215
         MaxLength       =   100
         TabIndex        =   5
         Top             =   945
         Width           =   4920
      End
      Begin VB.TextBox txtDatos 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   315
         Index           =   1
         Left            =   1215
         MaxLength       =   30
         TabIndex        =   4
         Top             =   585
         Width           =   4920
      End
      Begin MSComCtl2.DTPicker txtfecha 
         Height          =   330
         Left            =   1215
         TabIndex        =   6
         Top             =   1305
         Width           =   1620
         _ExtentX        =   2858
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
         CalendarTitleBackColor=   12632256
         Format          =   16449537
         CurrentDate     =   38002
      End
      Begin MSComCtl2.DTPicker txtbaja 
         Height          =   330
         Left            =   4545
         TabIndex        =   7
         Top             =   1305
         Width           =   1620
         _ExtentX        =   2858
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
         CalendarTitleBackColor=   12632256
         Format          =   16449537
         CurrentDate     =   38002
      End
      Begin MSDataListLib.DataCombo cmbTipo 
         Height          =   315
         Left            =   1215
         TabIndex        =   3
         Top             =   225
         Width           =   4905
         _ExtentX        =   8652
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         Style           =   2
         Text            =   ""
      End
      Begin VB.Label Label1 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Tipo"
         Height          =   240
         Left            =   180
         TabIndex        =   30
         Top             =   270
         Width           =   420
      End
      Begin VB.Label lblCampos 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Importe"
         Height          =   240
         Index           =   1
         Left            =   180
         TabIndex        =   23
         Top             =   1740
         Width           =   600
      End
      Begin VB.Label lblCampos 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Fecha Baja"
         Height          =   240
         Index           =   0
         Left            =   3600
         TabIndex        =   22
         Top             =   1350
         Width           =   825
      End
      Begin VB.Label lblCampos 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Descripción"
         Height          =   240
         Index           =   4
         Left            =   180
         TabIndex        =   21
         Top             =   990
         Width           =   1185
      End
      Begin VB.Label lblCampos 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Fecha Pedido"
         Height          =   240
         Index           =   2
         Left            =   180
         TabIndex        =   20
         Top             =   1350
         Width           =   1095
      End
      Begin VB.Label lblCampos 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Código"
         Height          =   240
         Index           =   3
         Left            =   180
         TabIndex        =   19
         Top             =   630
         Width           =   945
      End
   End
   Begin MSComctlLib.ListView lista 
      Height          =   5220
      Left            =   30
      TabIndex        =   17
      Top             =   1290
      Width           =   13890
      _ExtentX        =   24500
      _ExtentY        =   9208
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
   Begin VB.Image imagen 
      Height          =   480
      Left            =   13320
      Picture         =   "frmClientes_Pedidos.frx":4C8E
      Top             =   0
      Width           =   480
   End
   Begin VB.Label lbltitulo 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Caption         =   "Pedidos del Cliente"
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
      Height          =   285
      Left            =   120
      TabIndex        =   25
      Top             =   90
      Width           =   5100
   End
   Begin VB.Shape fondo 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   555
      Left            =   0
      Top             =   0
      Width           =   13920
   End
End
Attribute VB_Name = "frmClientes_Pedidos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public PK As Long

Private Sub chkvigor_Click()
    cargar_datos
End Sub

Private Sub cmbFiltroTipo_Change()
    cargar_lista
End Sub
Private Sub cmdAdjuntar_Click()
    If lista.ListItems.Count = 0 Then Exit Sub
    With frmAdjuntos
        .TOBJETO = TOBJETO.TOBJETO_CLIENTES_PEDIDOS
        .COBJETO = lista.ListItems(lista.selectedItem.Index).Text
        .Show 1
    End With
    Set frmAdjuntos = Nothing
    Call PresentarDatos_DocumentosAdjuntos
End Sub

Private Sub cmdAnadir_Click()
    On Error GoTo fallo
    If validar = True Then
        Dim ocliente_pedido As New clsClientes_pedidos
        With ocliente_pedido
            .setCLIENTE_ID = PK
            .setTIPO_ID = cmbTipo.BoundText
            .setCODIGO = txtDatos(1)
            .setDESCRIPCION = txtDatos(2)
            .setFECHA_PEDIDO = Format(txtFecha.value, "yyyy-mm-dd")
            .setFECHA_BAJA = Format(txtbaja.value, "yyyy-mm-dd")
            .setIMPORTE = Replace(Format(txtDatos(0), "0.00"), ",", ".")
            .setREVISADO = chkRevisadoConforme.value
            .Insertar
        End With
        borrar_campos
        cargar_datos
        MsgBox "Pedido de cliente insertado correctamente.", vbInformation, App.Title
    End If
    Exit Sub
fallo:
    error_grave ("Error al insertar el pedido : " & Err.Description)
End Sub
Private Sub cmdDetalle_Click()
    If lista.ListItems.Count > 0 Then
        frmClientes_Detalle_Pedido.PK = lista.ListItems(lista.selectedItem.Index).Text
        frmClientes_Detalle_Pedido.Show 1
    End If
End Sub

Private Sub cmdEliminar_Click()
    If lista.ListItems.Count > 0 Then
        If MsgBox("¿Seguro de eliminar el pedido : " & lista.ListItems(lista.selectedItem.Index).SubItems(1) & "?", vbQuestion + vbYesNo, App.Title) = vbYes Then
            Dim ocliente_pedido As New clsClientes_pedidos
            ocliente_pedido.Eliminar (lista.ListItems(lista.selectedItem.Index))
            'M1539-I
            listaAdjuntos.ListItems.Clear
            'M1539-F
            borrar_campos
            cargar_datos
            MsgBox "Pedido de cliente eliminado correctamente.", vbInformation, App.Title
        End If
    End If
End Sub

Private Sub cmdModificar_Click()
    If lista.ListItems.Count = 0 Then
        Exit Sub
    End If
    On Error GoTo fallo
    If validar = True Then
        Dim ocliente_pedido As New clsClientes_pedidos
        With ocliente_pedido
            .setCLIENTE_ID = PK
            .setTIPO_ID = cmbTipo.BoundText
            .setCODIGO = txtDatos(1)
            .setDESCRIPCION = txtDatos(2)
            .setFECHA_PEDIDO = Format(txtFecha.value, "yyyy-mm-dd")
            .setFECHA_BAJA = Format(txtbaja.value, "yyyy-mm-dd")
            .setIMPORTE = Replace(Format(txtDatos(0), "0.00"), ",", ".")
            .setREVISADO = chkRevisadoConforme.value
            .Modificar (lista.ListItems(lista.selectedItem.Index))
        End With
        borrar_campos
        cargar_datos
        MsgBox "Pedido de cliente modificado correctamente.", vbInformation, App.Title
    End If
    Exit Sub
fallo:
    error_grave ("Error al modificar el pedido : " & Err.Description)
End Sub

Private Sub cmdSalir_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    log (Me.Name)
    cargar_botones Me
    txtFecha = Date
    txtbaja = Date
    Dim oDeco As New clsDecodificadora
    oDeco.cargar_combo cmbTipo, DECODIFICADORA.PEDIDOS_CLIENTES_TIPOS
    oDeco.cargar_combo cmbFiltroTipo, DECODIFICADORA.PEDIDOS_CLIENTES_TIPOS
    Call cabecera
    If PK <> 0 Then
        cargar_datos
    End If
End Sub

Public Sub cabecera()
    With lista.ColumnHeaders
        .Add , , "ID", 1, lvwColumnLeft
        .Add , , "Tipo", 2000, lvwColumnLeft
        .Add , , "Código", 2500, lvwColumnCenter
        .Add , , "Descripción", 3500, lvwColumnLeft
        .Add , , "F.Alta", 1050, lvwColumnCenter
        .Add , , "F.Pedido", 1050, lvwColumnCenter
        .Add , , "F.Baja", 1050, lvwColumnCenter
        .Add , , "Importe", 1050, lvwColumnRight
        .Add , , "ID_TIPO", 1, lvwColumnRight
        .Add , , "Restan", 1200, lvwColumnRight
        .Add , , "Revisado", 1, lvwColumnCenter
    End With
    
    With listaAdjuntos.ColumnHeaders
        .Add , , "ID", 1, lvwColumnLeft
        .Add , , "Documento", listaAdjuntos.Width - 1200, lvwColumnLeft
        .Add , , "Fecha", 1200, lvwColumnCenter
    End With
End Sub
Private Sub lista_Click()
    listaAdjuntos.ListItems.Clear
    If lista.ListItems.Count > 0 Then
        txtDatos(1) = lista.ListItems(lista.selectedItem.Index).SubItems(2)
        txtDatos(2) = lista.ListItems(lista.selectedItem.Index).SubItems(3)
        txtDatos(0) = lista.ListItems(lista.selectedItem.Index).SubItems(7)
        txtFecha = lista.ListItems(lista.selectedItem.Index).SubItems(5)
        txtbaja = lista.ListItems(lista.selectedItem.Index).SubItems(6)
        cmbTipo.BoundText = lista.ListItems(lista.selectedItem.Index).SubItems(8)
        PresentarDatos_DocumentosAdjuntos
        chkRevisadoConforme.value = lista.ListItems(lista.selectedItem.Index).SubItems(10)
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
    If lista.ListItems.Count > 0 Then
        cmdDetalle_Click
    End If
End Sub

Private Sub listaAdjuntos_DblClick()
    If listaAdjuntos.ListItems.Count = 0 Then
        MsgBox "Seleccione algún archivo de la lista.", vbExclamation, App.Title
        Exit Sub
    End If
    Dim oAdjunto As New clsAdjuntos
    oAdjunto.CargarDocumento TOBJETO.TOBJETO_CLIENTES_PEDIDOS, lista.ListItems(lista.selectedItem.Index).Text, 0, listaAdjuntos.ListItems(listaAdjuntos.selectedItem.Index).Text, True
    Set oAdjunto = Nothing

End Sub

Private Sub txtdatos_GotFocus(Index As Integer)
    txtDatos(Index).BackColor = &H80C0FF
    txtDatos(Index).SelStart = 0
    txtDatos(Index).SelLength = Len(txtDatos(Index))
End Sub

Private Sub txtDatos_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys "{Tab}", True
        KeyAscii = 0 ' Para evitar el "bip" del sistema
    End If
    
    If KeyAscii = 46 And Index = 0 Then
         KeyAscii = 44
    End If
End Sub
Private Sub txtdatos_LostFocus(Index As Integer)
    txtDatos(Index).BackColor = vbWhite
End Sub
Private Sub cargar_datos()
    Dim oCliente As New clsCliente
    With oCliente
        If .CargaCliente(PK) = True Then
            lbltitulo = "Pedidos del cliente : " & .getNOMBRE
            cargar_lista
        End If
    End With
    Set oCliente = Nothing
End Sub
Private Sub cargar_lista()
    Dim oCliente_Pedidos As New clsClientes_pedidos
    Dim rs As ADODB.Recordset
    If chkvigor.value = False Then
       Set rs = oCliente_Pedidos.Listado(PK, cmbFiltroTipo.BoundText, txtFiltro(0), txtFiltro(1))
    Else
       Set rs = oCliente_Pedidos.Listado_Vigor(PK, cmbFiltroTipo.BoundText, txtFiltro(0), txtFiltro(1))
    End If
    Dim oDoc As New clsDocs_pago
    lista.ListItems.Clear
    If rs.RecordCount <> 0 Then
        Do
            With lista.ListItems.Add(, , rs(0))
                .SubItems(1) = rs(6) ' TIPO
                .SubItems(2) = rs(1)
                .SubItems(3) = rs(2)
                .SubItems(4) = Format(rs(8), "dd-mm-yyyy") ' F.Alta
                .SubItems(5) = Format(rs(3), "dd-mm-yyyy") ' F.Pedido
                .SubItems(6) = Format(rs(4), "dd-mm-yyyy") ' F.Baja
                .SubItems(7) = Format(rs(5), "currency")
                .SubItems(8) = rs(7) ' TIPO_ID
                ' Pendiente
'                .SubItems(9) = moneda(rs(5) - oDoc.Documentos_por_pedido_suma_importe(rs(0)))
                .SubItems(9) = moneda(rs(9))
                .SubItems(10) = rs(10)
            End With
            rs.MoveNext
        Loop Until rs.EOF
        lista_Click
    End If
    Set oCliente_Pedidos = Nothing
    PresentarDatos_DocumentosAdjuntos
End Sub
Public Function validar() As Boolean
    validar = True
    If cmbTipo.Text = "" Then
        MsgBox "Debe introducir el tipo de pedido.", vbInformation, App.Title
        validar = False
        Exit Function
    End If
    If Trim(txtDatos(0)) = "" Then
        MsgBox "Debe introducir un importe para el pedido.", vbInformation, App.Title
        validar = False
        Exit Function
    End If
    If Trim(txtDatos(1)) = "" Then
        MsgBox "Debe introducir un código de pedido.", vbInformation, App.Title
        validar = False
        Exit Function
    End If
    If Trim(txtDatos(2)) = "" Then
        MsgBox "Debe introducir una descripción para el pedido.", vbInformation, App.Title
        validar = False
        Exit Function
    End If
    If chkRevisadoConforme.value = Unchecked Then
        MsgBox "Debe marcar el pedido como Revisado y Conforme.", vbExclamation, App.Title
        validar = False
        Exit Function
    End If
End Function
Public Sub borrar_campos()
    txtDatos(0) = ""
    txtDatos(1) = ""
    txtDatos(2) = ""
    txtFecha = Date
    txtbaja = Date
End Sub
Private Sub PresentarDatos_DocumentosAdjuntos()
    If lista.ListItems.Count = 0 Then Exit Sub
    listaAdjuntos.ListItems.Clear
    Dim oAdjunto As New clsAdjuntos
    Dim rs As ADODB.Recordset
    Set rs = oAdjunto.Listado(TOBJETO.TOBJETO_CLIENTES_PEDIDOS, lista.ListItems(lista.selectedItem.Index).Text, "", "")
    If rs.RecordCount > 0 Then
        Do
            With listaAdjuntos.ListItems.Add(, , rs(0))
                 .SubItems(1) = rs(2)
                 .SubItems(2) = Format(rs(3), "dd-mm-yyyy")
            End With
            rs.MoveNext
        Loop Until rs.EOF
    End If
End Sub

Private Sub txtfiltro_Change(Index As Integer)
    cargar_lista
End Sub
